' ShareRibbon\Ribbon\BaseOfficeRibbon.Designer.vb
Partial Class BaseOfficeRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Sub New(ByVal factory As Microsoft.Office.Tools.Ribbon.RibbonFactory)
        MyBase.New(factory)
        InitializeComponent()
    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub InitializeComponent()
        Me.TabAI = Me.Factory.CreateRibbonTab
        Me.GroupAI = Me.Factory.CreateRibbonGroup
        Me.ConfigApiButton = Me.Factory.CreateRibbonButton
        Me.DataAnalysisButton = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.PromptConfigButton = Me.Factory.CreateRibbonButton
        Me.ChatButton = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.AboutButton = Me.Factory.CreateRibbonButton
        Me.ClearCacheButton = Me.Factory.CreateRibbonButton

        ' 新增：Deepseek 专用 Group
        Me.GroupDeepseek = Me.Factory.CreateRibbonGroup
        Me.DeepseekButton = Me.Factory.CreateRibbonButton()

        ' 第二个Group和按钮
        Me.GroupTools = Me.Factory.CreateRibbonGroup
        Me.WebCaptureButton = Me.Factory.CreateRibbonButton
        Me.SpotlightButton = Me.Factory.CreateRibbonButton

        Me.DeepseekButton = Me.Factory.CreateRibbonButton()
        Me.BatchDataGenButton = Me.Factory.CreateRibbonButton()
        Me.MCPButton = Me.Factory.CreateRibbonButton()


        Me.TabAI.SuspendLayout()
        Me.GroupAI.SuspendLayout()
        Me.SuspendLayout()

        ' TabAI
        Me.TabAI.Groups.Add(Me.GroupDeepseek)  ' 首先添加Deepseek Group
        Me.TabAI.Groups.Add(Me.GroupAI)
        Me.TabAI.Groups.Add(Me.GroupTools)

        Me.TabAI.Label = "AI助手"
        Me.TabAI.Name = "TabAI"

        ' GroupDeepseek - 新的Deepseek专用Group
        Me.GroupDeepseek.Items.Add(Me.DeepseekButton)
        Me.GroupDeepseek.Label = "免费强化版"
        Me.GroupDeepseek.Name = "GroupDeepseek"

        ' 配置Deepseek按钮 - 设置为大图标
        Me.DeepseekButton.Label = "Deepseek"
        Me.DeepseekButton.Name = "DeepseekButton"
        Me.DeepseekButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeepseekButton.ShowImage = True
        Me.DeepseekButton.ScreenTip = "免费增强版"
        Me.DeepseekButton.SuperTip = "在原有对话基础上，增加Agent执行能力"


        ' GroupAI
        Me.GroupAI.Items.Add(Me.ConfigApiButton)
        Me.GroupAI.Items.Add(Me.DataAnalysisButton)
        Me.GroupAI.Items.Add(Me.Separator1)
        Me.GroupAI.Items.Add(Me.PromptConfigButton)
        Me.GroupAI.Items.Add(Me.ChatButton)
        Me.GroupAI.Items.Add(Me.Separator2)
        Me.GroupAI.Items.Add(Me.AboutButton)
        Me.GroupAI.Items.Add(Me.ClearCacheButton)
        Me.GroupAI.Label = "AI大模型"
        Me.GroupAI.Name = "GroupAI"

        ' ConfigApiButton
        Me.ConfigApiButton.Label = "配置API"
        Me.ConfigApiButton.Name = "ConfigApiButton"
        Me.ConfigApiButton.ShowImage = True
        Me.ConfigApiButton.SuperTip = "使用AI功能前需要配置apiKey"

        ' DataAnalysisButton
        Me.DataAnalysisButton.Label = "数据分析"
        Me.DataAnalysisButton.Name = "DataAnalysisButton"
        Me.DataAnalysisButton.ShowImage = True

        ' PromptConfigButton
        Me.PromptConfigButton.Label = "提示词"
        Me.PromptConfigButton.Name = "PromptConfigButton"
        Me.PromptConfigButton.ShowImage = True

        ' ChatButton
        Me.ChatButton.Label = "Chat AI"
        Me.ChatButton.Name = "ChatButton"
        Me.ChatButton.ShowImage = True

        ' AboutButton
        Me.AboutButton.Label = "关于"
        Me.AboutButton.Name = "AboutButton"
        Me.AboutButton.ShowImage = True

        ' ClearCacheButton
        Me.ClearCacheButton.Label = "清理缓存"
        Me.ClearCacheButton.Name = "ClearCacheButton"
        Me.ClearCacheButton.ShowImage = True

        ' Separators
        Me.Separator1.Name = "Separator1"
        Me.Separator2.Name = "Separator2"

        ' 第二个Group
        ' GroupTools
        Me.GroupTools.Items.Add(Me.WebCaptureButton)
        Me.GroupTools.Items.Add(Me.SpotlightButton)
        Me.GroupTools.Label = "工具箱"
        Me.GroupTools.Name = "GroupTools"

        ' WebCaptureButton
        Me.WebCaptureButton.Label = "抓取网页"
        Me.WebCaptureButton.Name = "WebCaptureButton"
        Me.WebCaptureButton.ShowImage = True
        Me.WebCaptureButton.SuperTip = "打开网页捕获工具"

        ' SpotlightButton
        Me.SpotlightButton.Label = "聚光灯"
        Me.SpotlightButton.Name = "SpotlightButton"
        Me.SpotlightButton.ShowImage = True
        Me.SpotlightButton.SuperTip = "高亮选中单元格所在的行和列"


        ' 配置批量数据生成按钮
        Me.BatchDataGenButton.Label = "批量数据生成"
        Me.BatchDataGenButton.Name = "BatchDataGenButton"
        Me.BatchDataGenButton.ScreenTip = "配置和生成批量数据"
        Me.BatchDataGenButton.SuperTip = "配置字段、列关系并生成数据到工作簿"
        Me.GroupTools.Items.Add(Me.BatchDataGenButton)

        ' 配置MCP按钮
        Me.MCPButton.Label = "MCP"
        Me.MCPButton.Name = "MCPButton"
        Me.MCPButton.ScreenTip = "MCP服务器配置"
        Me.MCPButton.SuperTip = "配置MCP服务器并作为客户端调用大模型"
        Me.GroupTools.Items.Add(Me.MCPButton)


        ' BaseOfficeRibbon
        Me.Name = "BaseOfficeRibbon"
        Me.Tabs.Add(Me.TabAI)

        Me.TabAI.ResumeLayout(False)
        Me.TabAI.PerformLayout()
        Me.GroupAI.ResumeLayout(False)
        Me.GroupAI.PerformLayout()
        Me.ResumeLayout(False)
    End Sub

    Protected WithEvents TabAI As Microsoft.Office.Tools.Ribbon.RibbonTab
    Protected WithEvents GroupAI As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ConfigApiButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents DataAnalysisButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents PromptConfigButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ChatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents AboutButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ClearCacheButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Protected WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator

    ' 新增：Deepseek 专用 Group 声明
    Protected WithEvents GroupDeepseek As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents DeepseekButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' 在 Class BaseOfficeRibbon 的底部添加这些控件声明
    Protected WithEvents GroupTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents WebCaptureButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents SpotlightButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    Protected WithEvents BatchDataGenButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents MCPButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class