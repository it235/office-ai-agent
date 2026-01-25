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

    ' 关于我按钮点击事件
    Private Sub AboutButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AboutButton.Click
        MsgBox("大家好，我是B站的君哥，账号 君哥聊编程 。该插件的灵感是来自于一位B站的粉丝，他是银行审计相关的工作，经常与表格打交道，很多时候表格中的数据无法通过固定的公式来计算，但是在人类理解上又具有相同的意义，所以Excel AI诞生了。
插件在持续优化中，我本身与Excel打交道比较少，如果你有更多好的idea可以过来给我留言或评论，不断完善该插件。ExcelAi数据的默认存放目录在当前用户/文档/" + ConfigSettings.OfficeAiAppDataFolder + "下。")
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
End Class