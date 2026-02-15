' ShareRibbon\Ribbon\BaseOfficeRibbon.vb
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Excel
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json.Linq


Public MustInherit Class BaseOfficeRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim apiConfig As New ConfigManager()
        apiConfig.LoadConfig()
        Dim promptConfig As New ConfigPromptForm(GetApplication())
        promptConfig.LoadConfig()
        InitializeBaseRibbon()
    End Sub

    Protected Overridable Sub InitializeBaseRibbon()
        ' 基类初始化方法，子类可以重写
    End Sub

    ' 关于我按钮点击事件 - 显示带git链接的对话框
    Private Sub AboutButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AboutButton.Click
        Dim aboutForm As New AboutForm()
        aboutForm.ShowDialog()
    End Sub

    ' 清理缓存配置按钮点击事件
    Private Sub ClearCacheConfig_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ClearCacheButton.Click
        ' 弹出确认框
        Dim result = MessageBox.Show("将彻底删除‘文档\" & ConfigSettings.OfficeAiAppDataFolder & "’目录下所有的配置，历史聊天记录信息，清理后不可恢复，您确定要清理吗？", "确认操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result <> DialogResult.OK Then
            Return
        End If

        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & ConfigSettings.OfficeAiAppDataFolder
        If System.IO.Directory.Exists(appDataPath) Then
            Try
                Dim files As String() = System.IO.Directory.GetFiles(appDataPath)
                For Each file In files
                    System.IO.File.Delete(file)
                Next
                MsgBox("缓存配置已清理，请重启Office相关应用！")
            Catch ex As Exception
                MsgBox("清理缓存配置时出错：" & ex.Message, vbCritical)
            End Try
        Else
            MsgBox("缓存目录不存在！")
        End If
    End Sub

    ' 点击Ribbon区的配置API按钮后触发
    Private Sub ConfigApiButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfigApiButton.Click
        ' 创建并显示配置 API 的对话框
        Dim configForm As New ConfigApiForm()
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub
    Private Sub PromptConfigButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PromptConfigButton.Click
        ' 创建并显示配置 API 的对话框
        Dim configForm As New ConfigPromptForm(GetApplication())
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub

    ' 教学文档按钮点击事件 - 根据应用类型跳转不同URL
    Private Sub StudyButton_Click(sender As Object, e As RibbonControlEventArgs) Handles StudyButton.Click
        Dim appInfo = GetApplication()
        Dim url As String = "https://www.officeso.cn/study/"

        Select Case appInfo.Type
            Case OfficeApplicationType.Word
                url &= "word"
            Case OfficeApplicationType.Excel
                url &= "excel"
            Case OfficeApplicationType.PowerPoint
                url &= "ppt"
            Case Else
                url &= "word"
        End Select

        Try
            System.Diagnostics.Process.Start(url)
        Catch ex As Exception
            MessageBox.Show("无法打开教学文档链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' 定义 ComboBoxItem 类
    Private Class ComboBoxItem
        Public Property Text As String
        Public Property Value As String

        Public Sub New(text As String, value As String)
            Me.Text = text
            Me.Value = value
        End Sub

        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class

    ' AI聊天实现
    Protected MustOverride Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ChatButton.Click

    ' web爬虫实现
    Protected MustOverride Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs) Handles WebCaptureButton.Click

    ' 聚光灯实现（跟随鼠标选中整行和整列并高亮）
    Protected MustOverride Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SpotlightButton.Click

    ' 数据魔法分析实现
    Protected MustOverride Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    Protected MustOverride Function GetApplication() As ApplicationInfo


    ' 新增：校对与排版按钮的抽象事件（由子类实现具体流程）
    Protected MustOverride Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ProofreadButton.Click
    Protected MustOverride Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ReformatButton.Click

    ' Deepseek按钮点击事件
    Protected MustOverride Sub DeepseekButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DeepseekButton.Click

    ' Doubao按钮点击事件
    Protected MustOverride Sub DoubaoButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DoubaoButton.Click

    ' 批量数据生成按钮点击事件
    Protected MustOverride Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs) Handles BatchDataGenButton.Click

    ' MCP按钮点击事件
    Protected MustOverride Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles MCPButton.Click

    ' 一键翻译按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TranslateButton.Click

    ' AI续写按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ContinuationButton.Click

    ' 模板排版按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TemplateFormatButton.Click
End Class