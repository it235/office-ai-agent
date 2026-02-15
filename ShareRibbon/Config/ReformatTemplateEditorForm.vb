' ShareRibbon\Config\ReformatTemplateEditorForm.vb
' 排版模板编辑器窗体

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 实时预览回调委托 - 用于在Word中预览样式效果
''' </summary>
''' <param name="fontName">中文字体名</param>
''' <param name="fontSize">字号</param>
''' <param name="bold">是否加粗</param>
''' <param name="alignment">对齐方式</param>
''' <param name="firstLineIndent">首行缩进（字符数）</param>
''' <param name="lineSpacing">行距（倍数）</param>
Public Delegate Sub PreviewStyleCallback(fontName As String, fontSize As Double, bold As Boolean, alignment As String, firstLineIndent As Double, lineSpacing As Double)

''' <summary>
''' 模板占位符预览回调委托 - 用于实时预览占位符内容
''' </summary>
''' <param name="placeholderId">占位符ID</param>
''' <param name="content">占位符内容</param>
''' <param name="fontConfig">字体配置</param>
''' <param name="paragraphConfig">段落配置</param>
''' <param name="colorConfig">颜色配置</param>
Public Delegate Sub TemplatePlaceholderPreviewCallback(placeholderId As String, content As String, fontConfig As FontConfig, paragraphConfig As ParagraphConfig, colorConfig As ColorConfig)

''' <summary>
''' 排版模板编辑器窗体
''' </summary>
Public Class ReformatTemplateEditorForm
    Inherits Form

    Private _template As ReformatTemplate
    Private _isNewTemplate As Boolean
    
    ' 实时预览回调
    Private _previewCallback As PreviewStyleCallback

    ' 基本信息控件
    Private txtName As TextBox
    Private txtDescription As TextBox
    Private cboCategory As ComboBox
    Private txtAiGuidance As TextBox

    ' Tab控件
    Private tabControl As TabControl
    Private tabBasicInfo As TabPage
    Private tabLayout As TabPage
    Private tabBodyStyles As TabPage
    Private tabPageSettings As TabPage

    ' 版式列表
    Private lstLayoutElements As ListBox
    Private btnAddElement As Button
    Private btnRemoveElement As Button
    Private btnMoveUp As Button
    Private btnMoveDown As Button

    ' 版式元素编辑区
    Private pnlElementEdit As Panel
    Private txtElementName As TextBox
    Private txtElementDefaultValue As TextBox
    Private cboElementType As ComboBox
    Private cboElementFontCN As ComboBox
    Private cboElementFontSize As ComboBox
    Private chkElementBold As CheckBox
    Private cboElementAlignment As ComboBox

    ' 正文样式列表
    Private lstBodyStyles As ListBox
    Private btnAddStyle As Button
    Private btnRemoveStyle As Button

    ' 正文样式编辑区
    Private pnlStyleEdit As Panel
    Private txtStyleName As TextBox
    Private txtStyleCondition As TextBox
    Private cboStyleFontCN As ComboBox
    Private cboStyleFontSize As ComboBox
    Private chkStyleBold As CheckBox
    Private cboStyleAlignment As ComboBox
    Private numStyleFirstIndent As NumericUpDown
    Private numStyleLineSpacing As NumericUpDown

    ' 页面设置控件
    Private numMarginTop As NumericUpDown
    Private numMarginBottom As NumericUpDown
    Private numMarginLeft As NumericUpDown
    Private numMarginRight As NumericUpDown
    Private chkPageNumber As CheckBox
    Private txtPageNumberFormat As TextBox
    Private cboPageNumberPosition As ComboBox

    ' 底部按钮
    Private btnSave As Button
    Private btnCancel As Button

    Public Sub New(Optional template As ReformatTemplate = Nothing, Optional previewCallback As PreviewStyleCallback = Nothing)
        _isNewTemplate = template Is Nothing
        _template = If(template, New ReformatTemplate())
        _previewCallback = previewCallback

        InitializeForm()
        LoadTemplateData()
    End Sub

    Private Sub InitializeForm()
        Me.Text = If(_isNewTemplate, "新建排版模板", "编辑排版模板")
        Me.Size = New Size(700, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 创建Tab控件
        tabControl = New TabControl With {
            .Location = New Point(10, 10),
            .Size = New Size(665, 500),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }

        ' 基本信息Tab
        tabBasicInfo = New TabPage("基本信息")
        CreateBasicInfoTab()
        tabControl.TabPages.Add(tabBasicInfo)

        ' 版式Tab
        tabLayout = New TabPage("版式配置")
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

        Me.Controls.Add(tabControl)

        ' 底部按钮
        btnSave = New Button With {
            .Text = "保存",
            .Location = New Point(500, 520),
            .Size = New Size(80, 30),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        Me.Controls.Add(btnSave)

        btnCancel = New Button With {
            .Text = "取消",
            .Location = New Point(590, 520),
            .Size = New Size(80, 30),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancel)
    End Sub

    Private Sub CreateBasicInfoTab()
        Dim y As Integer = 20

        ' 模板名称
        tabBasicInfo.Controls.Add(New Label With {.Text = "模板名称:", .Location = New Point(20, y), .AutoSize = True})
        txtName = New TextBox With {.Location = New Point(120, y - 3), .Size = New Size(300, 23)}
        tabBasicInfo.Controls.Add(txtName)
        y += 35

        ' 分类
        tabBasicInfo.Controls.Add(New Label With {.Text = "分类:", .Location = New Point(20, y), .AutoSize = True})
        cboCategory = New ComboBox With {
            .Location = New Point(120, y - 3),
            .Size = New Size(150, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboCategory.Items.AddRange({"通用", "行政", "学术", "商务"})
        tabBasicInfo.Controls.Add(cboCategory)
        y += 35

        ' 描述
        tabBasicInfo.Controls.Add(New Label With {.Text = "描述:", .Location = New Point(20, y), .AutoSize = True})
        txtDescription = New TextBox With {
            .Location = New Point(120, y - 3),
            .Size = New Size(500, 60),
            .Multiline = True
        }
        tabBasicInfo.Controls.Add(txtDescription)
        y += 75

        ' AI说明
        tabBasicInfo.Controls.Add(New Label With {.Text = "AI说明:", .Location = New Point(20, y), .AutoSize = True})
        txtAiGuidance = New TextBox With {
            .Location = New Point(120, y - 3),
            .Size = New Size(500, 100),
            .Multiline = True
        }
        tabBasicInfo.Controls.Add(txtAiGuidance)

        ' 提示文本
        Dim tipLabel = New Label With {
            .Text = "AI说明：提供给AI的额外上下文信息，帮助AI更好地理解排版需求。",
            .Location = New Point(120, y + 105),
            .AutoSize = True,
            .ForeColor = Color.Gray,
            .Font = New Font(Me.Font.FontFamily, 8)
        }
        tabBasicInfo.Controls.Add(tipLabel)
    End Sub

    Private Sub CreateLayoutTab()
        ' 左侧列表
        tabLayout.Controls.Add(New Label With {.Text = "骨架元素:", .Location = New Point(20, 15), .AutoSize = True})

        lstLayoutElements = New ListBox With {
            .Location = New Point(20, 35),
            .Size = New Size(200, 350)
        }
        AddHandler lstLayoutElements.SelectedIndexChanged, AddressOf LstLayoutElements_SelectedIndexChanged
        tabLayout.Controls.Add(lstLayoutElements)

        ' 列表按钮
        btnAddElement = New Button With {.Text = "+", .Location = New Point(20, 390), .Size = New Size(40, 25)}
        AddHandler btnAddElement.Click, AddressOf BtnAddElement_Click
        tabLayout.Controls.Add(btnAddElement)

        btnRemoveElement = New Button With {.Text = "-", .Location = New Point(65, 390), .Size = New Size(40, 25)}
        AddHandler btnRemoveElement.Click, AddressOf BtnRemoveElement_Click
        tabLayout.Controls.Add(btnRemoveElement)

        btnMoveUp = New Button With {.Text = "↑", .Location = New Point(130, 390), .Size = New Size(40, 25)}
        AddHandler btnMoveUp.Click, AddressOf BtnMoveUp_Click
        tabLayout.Controls.Add(btnMoveUp)

        btnMoveDown = New Button With {.Text = "↓", .Location = New Point(175, 390), .Size = New Size(40, 25)}
        AddHandler btnMoveDown.Click, AddressOf BtnMoveDown_Click
        tabLayout.Controls.Add(btnMoveDown)

        ' 右侧编辑区
        pnlElementEdit = New Panel With {
            .Location = New Point(240, 15),
            .Size = New Size(400, 420),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tabLayout.Controls.Add(pnlElementEdit)

        Dim y As Integer = 15

        pnlElementEdit.Controls.Add(New Label With {.Text = "元素名称:", .Location = New Point(15, y), .AutoSize = True})
        txtElementName = New TextBox With {.Location = New Point(100, y - 3), .Size = New Size(150, 23)}
        pnlElementEdit.Controls.Add(txtElementName)
        y += 35

        pnlElementEdit.Controls.Add(New Label With {.Text = "元素类型:", .Location = New Point(15, y), .AutoSize = True})
        cboElementType = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(150, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboElementType.Items.AddRange({"text", "redLine", "separator"})
        pnlElementEdit.Controls.Add(cboElementType)
        y += 35

        pnlElementEdit.Controls.Add(New Label With {.Text = "默认值:", .Location = New Point(15, y), .AutoSize = True})
        txtElementDefaultValue = New TextBox With {.Location = New Point(100, y - 3), .Size = New Size(280, 23)}
        pnlElementEdit.Controls.Add(txtElementDefaultValue)
        y += 35

        pnlElementEdit.Controls.Add(New Label With {.Text = "中文字体:", .Location = New Point(15, y), .AutoSize = True})
        cboElementFontCN = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(150, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }
        cboElementFontCN.Items.AddRange(GetCommonChineseFonts())
        cboElementFontCN.Text = "宋体"
        AddHandler cboElementFontCN.TextChanged, AddressOf OnElementStyleChanged
        AddHandler cboElementFontCN.SelectedIndexChanged, AddressOf OnElementStyleChanged
        pnlElementEdit.Controls.Add(cboElementFontCN)
        y += 35

        pnlElementEdit.Controls.Add(New Label With {.Text = "字号:", .Location = New Point(15, y), .AutoSize = True})
        cboElementFontSize = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(80, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }
        cboElementFontSize.Items.AddRange(GetCommonFontSizes())
        cboElementFontSize.Text = "12"
        AddHandler cboElementFontSize.TextChanged, AddressOf OnElementStyleChanged
        AddHandler cboElementFontSize.SelectedIndexChanged, AddressOf OnElementStyleChanged
        pnlElementEdit.Controls.Add(cboElementFontSize)

        chkElementBold = New CheckBox With {.Text = "加粗", .Location = New Point(200, y - 3), .AutoSize = True}
        AddHandler chkElementBold.CheckedChanged, AddressOf OnElementStyleChanged
        pnlElementEdit.Controls.Add(chkElementBold)
        y += 35

        pnlElementEdit.Controls.Add(New Label With {.Text = "对齐方式:", .Location = New Point(15, y), .AutoSize = True})
        cboElementAlignment = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(120, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboElementAlignment.Items.AddRange({"left", "center", "right", "justify"})
        AddHandler cboElementAlignment.SelectedIndexChanged, AddressOf OnElementStyleChanged
        pnlElementEdit.Controls.Add(cboElementAlignment)

        ' 保存元素按钮
        Dim btnSaveElement As New Button With {
            .Text = "保存元素",
            .Location = New Point(15, 350),
            .Size = New Size(100, 30)
        }
        AddHandler btnSaveElement.Click, AddressOf BtnSaveElement_Click
        pnlElementEdit.Controls.Add(btnSaveElement)
    End Sub

    Private Sub CreateBodyStylesTab()
        ' 左侧列表
        tabBodyStyles.Controls.Add(New Label With {.Text = "正文样式:", .Location = New Point(20, 15), .AutoSize = True})

        lstBodyStyles = New ListBox With {
            .Location = New Point(20, 35),
            .Size = New Size(200, 350)
        }
        AddHandler lstBodyStyles.SelectedIndexChanged, AddressOf LstBodyStyles_SelectedIndexChanged
        tabBodyStyles.Controls.Add(lstBodyStyles)

        ' 列表按钮
        btnAddStyle = New Button With {.Text = "+", .Location = New Point(20, 390), .Size = New Size(40, 25)}
        AddHandler btnAddStyle.Click, AddressOf BtnAddStyle_Click
        tabBodyStyles.Controls.Add(btnAddStyle)

        btnRemoveStyle = New Button With {.Text = "-", .Location = New Point(65, 390), .Size = New Size(40, 25)}
        AddHandler btnRemoveStyle.Click, AddressOf BtnRemoveStyle_Click
        tabBodyStyles.Controls.Add(btnRemoveStyle)

        ' 右侧编辑区
        pnlStyleEdit = New Panel With {
            .Location = New Point(240, 15),
            .Size = New Size(400, 420),
            .BorderStyle = BorderStyle.FixedSingle
        }
        tabBodyStyles.Controls.Add(pnlStyleEdit)

        Dim y As Integer = 15

        pnlStyleEdit.Controls.Add(New Label With {.Text = "样式名称:", .Location = New Point(15, y), .AutoSize = True})
        txtStyleName = New TextBox With {.Location = New Point(100, y - 3), .Size = New Size(150, 23)}
        pnlStyleEdit.Controls.Add(txtStyleName)
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "匹配条件:", .Location = New Point(15, y), .AutoSize = True})
        txtStyleCondition = New TextBox With {.Location = New Point(100, y - 3), .Size = New Size(280, 23)}
        pnlStyleEdit.Controls.Add(txtStyleCondition)
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "中文字体:", .Location = New Point(15, y), .AutoSize = True})
        cboStyleFontCN = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(150, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }
        cboStyleFontCN.Items.AddRange(GetCommonChineseFonts())
        cboStyleFontCN.Text = "宋体"
        AddHandler cboStyleFontCN.TextChanged, AddressOf OnBodyStyleChanged
        AddHandler cboStyleFontCN.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(cboStyleFontCN)
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "字号:", .Location = New Point(15, y), .AutoSize = True})
        cboStyleFontSize = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(80, 23),
            .DropDownStyle = ComboBoxStyle.DropDown
        }
        cboStyleFontSize.Items.AddRange(GetCommonFontSizes())
        cboStyleFontSize.Text = "12"
        AddHandler cboStyleFontSize.TextChanged, AddressOf OnBodyStyleChanged
        AddHandler cboStyleFontSize.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(cboStyleFontSize)

        chkStyleBold = New CheckBox With {.Text = "加粗", .Location = New Point(200, y - 3), .AutoSize = True}
        AddHandler chkStyleBold.CheckedChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(chkStyleBold)
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "对齐方式:", .Location = New Point(15, y), .AutoSize = True})
        cboStyleAlignment = New ComboBox With {
            .Location = New Point(100, y - 3),
            .Size = New Size(120, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboStyleAlignment.Items.AddRange({"left", "center", "right", "justify"})
        AddHandler cboStyleAlignment.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(cboStyleAlignment)
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "首行缩进:", .Location = New Point(15, y), .AutoSize = True})
        numStyleFirstIndent = New NumericUpDown With {
            .Location = New Point(100, y - 3),
            .Size = New Size(80, 23),
            .Minimum = 0,
            .Maximum = 10,
            .Value = 0,
            .DecimalPlaces = 1
        }
        AddHandler numStyleFirstIndent.ValueChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(numStyleFirstIndent)
        pnlStyleEdit.Controls.Add(New Label With {.Text = "字符", .Location = New Point(185, y), .AutoSize = True})
        y += 35

        pnlStyleEdit.Controls.Add(New Label With {.Text = "行距:", .Location = New Point(15, y), .AutoSize = True})
        numStyleLineSpacing = New NumericUpDown With {
            .Location = New Point(100, y - 3),
            .Size = New Size(80, 23),
            .Minimum = 1,
            .Maximum = 3,
            .Value = 1.5D,
            .DecimalPlaces = 1,
            .Increment = 0.1D
        }
        AddHandler numStyleLineSpacing.ValueChanged, AddressOf OnBodyStyleChanged
        pnlStyleEdit.Controls.Add(numStyleLineSpacing)
        pnlStyleEdit.Controls.Add(New Label With {.Text = "倍", .Location = New Point(185, y), .AutoSize = True})

        ' 保存样式按钮
        Dim btnSaveStyle As New Button With {
            .Text = "保存样式",
            .Location = New Point(15, 380),
            .Size = New Size(100, 30)
        }
        AddHandler btnSaveStyle.Click, AddressOf BtnSaveStyle_Click
        pnlStyleEdit.Controls.Add(btnSaveStyle)
    End Sub

    Private Sub CreatePageSettingsTab()
        Dim y As Integer = 30

        ' 页边距组
        Dim grpMargins As New GroupBox With {
            .Text = "页边距 (cm)",
            .Location = New Point(20, y),
            .Size = New Size(600, 100)
        }
        tabPageSettings.Controls.Add(grpMargins)

        grpMargins.Controls.Add(New Label With {.Text = "上:", .Location = New Point(30, 30), .AutoSize = True})
        numMarginTop = New NumericUpDown With {
            .Location = New Point(60, 27),
            .Size = New Size(70, 23),
            .Minimum = 0,
            .Maximum = 10,
            .Value = 2.54D,
            .DecimalPlaces = 2,
            .Increment = 0.1D
        }
        grpMargins.Controls.Add(numMarginTop)

        grpMargins.Controls.Add(New Label With {.Text = "下:", .Location = New Point(160, 30), .AutoSize = True})
        numMarginBottom = New NumericUpDown With {
            .Location = New Point(190, 27),
            .Size = New Size(70, 23),
            .Minimum = 0,
            .Maximum = 10,
            .Value = 2.54D,
            .DecimalPlaces = 2,
            .Increment = 0.1D
        }
        grpMargins.Controls.Add(numMarginBottom)

        grpMargins.Controls.Add(New Label With {.Text = "左:", .Location = New Point(290, 30), .AutoSize = True})
        numMarginLeft = New NumericUpDown With {
            .Location = New Point(320, 27),
            .Size = New Size(70, 23),
            .Minimum = 0,
            .Maximum = 10,
            .Value = 3.18D,
            .DecimalPlaces = 2,
            .Increment = 0.1D
        }
        grpMargins.Controls.Add(numMarginLeft)

        grpMargins.Controls.Add(New Label With {.Text = "右:", .Location = New Point(420, 30), .AutoSize = True})
        numMarginRight = New NumericUpDown With {
            .Location = New Point(450, 27),
            .Size = New Size(70, 23),
            .Minimum = 0,
            .Maximum = 10,
            .Value = 3.18D,
            .DecimalPlaces = 2,
            .Increment = 0.1D
        }
        grpMargins.Controls.Add(numMarginRight)

        y += 120

        ' 页码组
        Dim grpPageNumber As New GroupBox With {
            .Text = "页码设置",
            .Location = New Point(20, y),
            .Size = New Size(600, 100)
        }
        tabPageSettings.Controls.Add(grpPageNumber)

        chkPageNumber = New CheckBox With {.Text = "显示页码", .Location = New Point(30, 30), .AutoSize = True}
        grpPageNumber.Controls.Add(chkPageNumber)

        grpPageNumber.Controls.Add(New Label With {.Text = "格式:", .Location = New Point(30, 60), .AutoSize = True})
        txtPageNumberFormat = New TextBox With {
            .Location = New Point(70, 57),
            .Size = New Size(200, 23),
            .Text = "第{page}页 共{total}页"
        }
        grpPageNumber.Controls.Add(txtPageNumberFormat)

        grpPageNumber.Controls.Add(New Label With {.Text = "位置:", .Location = New Point(300, 60), .AutoSize = True})
        cboPageNumberPosition = New ComboBox With {
            .Location = New Point(340, 57),
            .Size = New Size(100, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboPageNumberPosition.Items.AddRange({"footer", "header"})
        cboPageNumberPosition.SelectedIndex = 0
        grpPageNumber.Controls.Add(cboPageNumberPosition)
    End Sub

    Private Sub LoadTemplateData()
        ' 基本信息
        txtName.Text = _template.Name
        txtDescription.Text = _template.Description
        cboCategory.SelectedItem = If(_template.Category, "通用")
        If cboCategory.SelectedIndex < 0 Then cboCategory.SelectedIndex = 0
        txtAiGuidance.Text = _template.AiGuidance

        ' 版式元素
        RefreshLayoutElementsList()

        ' 正文样式
        RefreshBodyStylesList()

        ' 页面设置
        If _template.PageSettings IsNot Nothing Then
            If _template.PageSettings.Margins IsNot Nothing Then
                numMarginTop.Value = CDec(_template.PageSettings.Margins.Top)
                numMarginBottom.Value = CDec(_template.PageSettings.Margins.Bottom)
                numMarginLeft.Value = CDec(_template.PageSettings.Margins.Left)
                numMarginRight.Value = CDec(_template.PageSettings.Margins.Right)
            End If
            If _template.PageSettings.PageNumber IsNot Nothing Then
                chkPageNumber.Checked = _template.PageSettings.PageNumber.Enabled
                txtPageNumberFormat.Text = _template.PageSettings.PageNumber.Format
                cboPageNumberPosition.SelectedItem = _template.PageSettings.PageNumber.Position
            End If
        End If
    End Sub

    Private Sub RefreshLayoutElementsList()
        lstLayoutElements.Items.Clear()
        If _template.Layout?.Elements IsNot Nothing Then
            For Each el In _template.Layout.Elements
                lstLayoutElements.Items.Add(el.Name)
            Next
        End If
    End Sub

    Private Sub RefreshBodyStylesList()
        lstBodyStyles.Items.Clear()
        If _template.BodyStyles IsNot Nothing Then
            For Each style In _template.BodyStyles
                lstBodyStyles.Items.Add(style.RuleName)
            Next
        End If
    End Sub

    Private Sub LstLayoutElements_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idx = lstLayoutElements.SelectedIndex
        If idx < 0 OrElse _template.Layout?.Elements Is Nothing OrElse idx >= _template.Layout.Elements.Count Then
            Return
        End If

        Dim el = _template.Layout.Elements(idx)
        txtElementName.Text = el.Name
        cboElementType.SelectedItem = el.ElementType
        txtElementDefaultValue.Text = el.DefaultValue
        cboElementFontCN.Text = If(el.Font?.FontNameCN, "宋体")
        cboElementFontSize.Text = If(el.Font?.FontSize, 12).ToString()
        chkElementBold.Checked = If(el.Font?.Bold, False)
        cboElementAlignment.SelectedItem = el.Paragraph?.Alignment
    End Sub

    Private Sub LstBodyStyles_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idx = lstBodyStyles.SelectedIndex
        If idx < 0 OrElse _template.BodyStyles Is Nothing OrElse idx >= _template.BodyStyles.Count Then
            Return
        End If

        Dim style = _template.BodyStyles(idx)
        txtStyleName.Text = style.RuleName
        txtStyleCondition.Text = style.MatchCondition
        cboStyleFontCN.Text = If(style.Font?.FontNameCN, "宋体")
        cboStyleFontSize.Text = If(style.Font?.FontSize, 12).ToString()
        chkStyleBold.Checked = If(style.Font?.Bold, False)
        cboStyleAlignment.SelectedItem = style.Paragraph?.Alignment
        numStyleFirstIndent.Value = CDec(If(style.Paragraph?.FirstLineIndent, 0))
        numStyleLineSpacing.Value = CDec(If(style.Paragraph?.LineSpacing, 1.5))
    End Sub

    Private Sub BtnAddElement_Click(sender As Object, e As EventArgs)
        If _template.Layout Is Nothing Then _template.Layout = New LayoutConfig()
        If _template.Layout.Elements Is Nothing Then _template.Layout.Elements = New List(Of LayoutElement)()

        Dim newElement As New LayoutElement With {
            .Name = "新元素",
            .ElementType = "text",
            .SortOrder = _template.Layout.Elements.Count + 1
        }
        _template.Layout.Elements.Add(newElement)
        RefreshLayoutElementsList()
        lstLayoutElements.SelectedIndex = lstLayoutElements.Items.Count - 1
    End Sub

    Private Sub BtnRemoveElement_Click(sender As Object, e As EventArgs)
        Dim idx = lstLayoutElements.SelectedIndex
        If idx < 0 Then Return

        _template.Layout.Elements.RemoveAt(idx)
        RefreshLayoutElementsList()
    End Sub

    Private Sub BtnMoveUp_Click(sender As Object, e As EventArgs)
        Dim idx = lstLayoutElements.SelectedIndex
        If idx <= 0 Then Return

        Dim el = _template.Layout.Elements(idx)
        _template.Layout.Elements.RemoveAt(idx)
        _template.Layout.Elements.Insert(idx - 1, el)
        RefreshLayoutElementsList()
        lstLayoutElements.SelectedIndex = idx - 1
    End Sub

    Private Sub BtnMoveDown_Click(sender As Object, e As EventArgs)
        Dim idx = lstLayoutElements.SelectedIndex
        If idx < 0 OrElse idx >= _template.Layout.Elements.Count - 1 Then Return

        Dim el = _template.Layout.Elements(idx)
        _template.Layout.Elements.RemoveAt(idx)
        _template.Layout.Elements.Insert(idx + 1, el)
        RefreshLayoutElementsList()
        lstLayoutElements.SelectedIndex = idx + 1
    End Sub

    Private Sub BtnSaveElement_Click(sender As Object, e As EventArgs)
        Dim idx = lstLayoutElements.SelectedIndex
        If idx < 0 Then Return

        Dim el = _template.Layout.Elements(idx)
        el.Name = txtElementName.Text
        el.ElementType = cboElementType.SelectedItem?.ToString()
        el.DefaultValue = txtElementDefaultValue.Text
        el.Font = New FontConfig With {
            .FontNameCN = cboElementFontCN.Text,
            .FontSize = ParseFontSize(cboElementFontSize.Text),
            .Bold = chkElementBold.Checked
        }
        el.Paragraph = New ParagraphConfig With {
            .Alignment = cboElementAlignment.SelectedItem?.ToString()
        }

        RefreshLayoutElementsList()
        lstLayoutElements.SelectedIndex = idx
    End Sub

    Private Sub BtnAddStyle_Click(sender As Object, e As EventArgs)
        If _template.BodyStyles Is Nothing Then _template.BodyStyles = New List(Of StyleRule)()

        Dim newStyle As New StyleRule With {
            .RuleName = "新样式",
            .SortOrder = _template.BodyStyles.Count + 1
        }
        _template.BodyStyles.Add(newStyle)
        RefreshBodyStylesList()
        lstBodyStyles.SelectedIndex = lstBodyStyles.Items.Count - 1
    End Sub

    Private Sub BtnRemoveStyle_Click(sender As Object, e As EventArgs)
        Dim idx = lstBodyStyles.SelectedIndex
        If idx < 0 Then Return

        _template.BodyStyles.RemoveAt(idx)
        RefreshBodyStylesList()
    End Sub

    Private Sub BtnSaveStyle_Click(sender As Object, e As EventArgs)
        Dim idx = lstBodyStyles.SelectedIndex
        If idx < 0 Then Return

        Dim style = _template.BodyStyles(idx)
        style.RuleName = txtStyleName.Text
        style.MatchCondition = txtStyleCondition.Text
        style.Font = New FontConfig With {
            .FontNameCN = cboStyleFontCN.Text,
            .FontSize = ParseFontSize(cboStyleFontSize.Text),
            .Bold = chkStyleBold.Checked
        }
        style.Paragraph = New ParagraphConfig With {
            .Alignment = cboStyleAlignment.SelectedItem?.ToString(),
            .FirstLineIndent = CDbl(numStyleFirstIndent.Value),
            .LineSpacing = CDbl(numStyleLineSpacing.Value)
        }

        RefreshBodyStylesList()
        lstBodyStyles.SelectedIndex = idx
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        ' 验证
        If String.IsNullOrWhiteSpace(txtName.Text) Then
            MessageBox.Show("请输入模板名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 保存基本信息
        _template.Name = txtName.Text.Trim()
        _template.Description = txtDescription.Text
        _template.Category = cboCategory.SelectedItem?.ToString()
        _template.AiGuidance = txtAiGuidance.Text

        ' 保存页面设置
        _template.PageSettings = New PageConfig With {
            .Margins = New MarginsConfig With {
                .Top = CDbl(numMarginTop.Value),
                .Bottom = CDbl(numMarginBottom.Value),
                .Left = CDbl(numMarginLeft.Value),
                .Right = CDbl(numMarginRight.Value)
            },
            .PageNumber = New PageNumberConfig With {
                .Enabled = chkPageNumber.Checked,
                .Format = txtPageNumberFormat.Text,
                .Position = cboPageNumberPosition.SelectedItem?.ToString()
            }
        }

        ' 保存到管理器
        If _isNewTemplate Then
            ReformatTemplateManager.Instance.AddTemplate(_template)
        Else
            ReformatTemplateManager.Instance.UpdateTemplate(_template)
        End If

        Me.DialogResult = DialogResult.OK
    End Sub
    
#Region "实时预览"
    
    ''' <summary>
    ''' 版式元素样式变化时触发预览
    ''' </summary>
    Private Sub OnElementStyleChanged(sender As Object, e As EventArgs)
        If _previewCallback Is Nothing Then Return
        
        Try
            Dim fontName = If(cboElementFontCN?.Text, "宋体")
            Dim fontSize = ParseFontSize(cboElementFontSize?.Text)
            Dim bold = If(chkElementBold IsNot Nothing, chkElementBold.Checked, False)
            Dim alignment = If(cboElementAlignment?.SelectedItem?.ToString(), "left")
            
            ' 版式元素无首行缩进和行距设置，使用默认值
            _previewCallback.Invoke(fontName, fontSize, bold, alignment, 0, 1.5)
        Catch ex As Exception
            Debug.WriteLine($"OnElementStyleChanged 预览出错: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 正文样式变化时触发预览
    ''' </summary>
    Private Sub OnBodyStyleChanged(sender As Object, e As EventArgs)
        If _previewCallback Is Nothing Then Return
        
        Try
            Dim fontName = If(cboStyleFontCN?.Text, "宋体")
            Dim fontSize = ParseFontSize(cboStyleFontSize?.Text)
            Dim bold = If(chkStyleBold IsNot Nothing, chkStyleBold.Checked, False)
            Dim alignment = If(cboStyleAlignment?.SelectedItem?.ToString(), "left")
            Dim firstIndent = If(numStyleFirstIndent IsNot Nothing, CDbl(numStyleFirstIndent.Value), 0)
            Dim lineSpacing = If(numStyleLineSpacing IsNot Nothing, CDbl(numStyleLineSpacing.Value), 1.5)
            
            _previewCallback.Invoke(fontName, fontSize, bold, alignment, firstIndent, lineSpacing)
        Catch ex As Exception
            Debug.WriteLine($"OnBodyStyleChanged 预览出错: {ex.Message}")
        End Try
    End Sub
    
#End Region

#Region "辅助方法"
    
    ''' <summary>
    ''' 获取常用中文字体列表
    ''' </summary>
    Private Shared Function GetCommonChineseFonts() As String()
        Return New String() {
            "宋体",
            "黑体",
            "楷体",
            "仿宋",
            "微软雅黑",
            "华文宋体",
            "华文黑体",
            "华文楷体",
            "华文仿宋",
            "华文中宋",
            "方正小标宋简体",
            "方正仿宋简体",
            "方正楷体简体",
            "方正黑体简体",
            "Times New Roman",
            "Arial"
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
        
        ' 中文字号映射表
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
                ' 尝试解析数字
                Dim result As Double
                If Double.TryParse(sizeText, result) Then
                    Return result
                Else
                    Return 12.0
                End If
        End Select
    End Function
    
#End Region
    
End Class
