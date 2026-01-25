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
        Me.DoubaoButton = Me.Factory.CreateRibbonButton()

        ' 新增：MCP 专用 Group
        Me.GroupMCP = Me.Factory.CreateRibbonGroup
        Me.MCPButton = Me.Factory.CreateRibbonButton()

        ' 第二个Group和按钮
        Me.GroupTools = Me.Factory.CreateRibbonGroup
        Me.WebCaptureButton = Me.Factory.CreateRibbonButton
        Me.SpotlightButton = Me.Factory.CreateRibbonButton

        Me.BatchDataGenButton = Me.Factory.CreateRibbonButton()

        ' 新增：一键翻译按钮
        Me.TranslateButton = Me.Factory.CreateRibbonButton

        ' 新增：校对与排版按钮
        Me.ProofreadButton = Me.Factory.CreateRibbonButton
        Me.ReformatButton = Me.Factory.CreateRibbonButton

        Me.TabAI.SuspendLayout()
        Me.GroupAI.SuspendLayout()
        Me.SuspendLayout()

        ' TabAI
        Me.TabAI.Groups.Add(Me.GroupDeepseek)  ' 首先添加Deepseek Group
        Me.TabAI.Groups.Add(Me.GroupAI)
        Me.TabAI.Groups.Add(Me.GroupTools)
        Me.TabAI.Groups.Add(Me.GroupMCP)

        Me.TabAI.Label = "AI助手"
        Me.TabAI.Name = "TabAI"

        ' GroupDeepseek - 新的Deepseek专用Group
        Me.GroupDeepseek.Items.Add(Me.DeepseekButton)
        Me.GroupDeepseek.Items.Add(Me.DoubaoButton)
        Me.GroupDeepseek.Label = "免费强化版"
        Me.GroupDeepseek.Name = "GroupDeepseek"

        ' 配置Deepseek按钮 - 设置为大图标
        Me.DeepseekButton.Label = "Deepseek"
        Me.DeepseekButton.Name = "DeepseekButton"
        Me.DeepseekButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeepseekButton.ShowImage = True
        Me.DeepseekButton.ScreenTip = "免费增强版"
        Me.DeepseekButton.SuperTip = "在原有对话基础上，增加Agent执行能力"

        ' 配置Doubao按钮 - 设置为大图标
        Me.DoubaoButton.Label = "Doubao"
        Me.DoubaoButton.Name = "DoubaoButton"
        Me.DoubaoButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DoubaoButton.ShowImage = True
        Me.DoubaoButton.ScreenTip = "豆包智能助手"
        Me.DoubaoButton.SuperTip = "基于豆包的智能对话助手，支持代码执行"

        ' GroupMCP - 新的MCP专用Group
        Me.GroupMCP.Items.Add(Me.MCPButton)
        Me.GroupMCP.Label = "MCP连接"
        Me.GroupMCP.Name = "GroupMCP"

        ' 配置MCP按钮 - 设置为大图标
        Me.MCPButton.Label = "MCP"
        Me.MCPButton.Name = "MCPButton"
        Me.MCPButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MCPButton.ShowImage = True
        Me.MCPButton.ScreenTip = "MCP服务器配置"
        Me.MCPButton.SuperTip = "配置MCP服务器并作为客户端调用大模型"

        ' GroupAI
        Me.GroupAI.Items.Add(Me.ConfigApiButton)
        Me.GroupAI.Items.Add(Me.DataAnalysisButton)
        Me.GroupAI.Items.Add(Me.Separator1)
        Me.GroupAI.Items.Add(Me.PromptConfigButton)
        Me.GroupAI.Items.Add(Me.ChatButton)
        Me.GroupAI.Items.Add(Me.TranslateButton)
        Me.GroupAI.Items.Add(Me.ProofreadButton)
        Me.GroupAI.Items.Add(Me.ReformatButton)
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
        ' 校对按钮
        Me.ProofreadButton.Label = "校对"
        Me.ProofreadButton.Name = "ProofreadButton"
        Me.ProofreadButton.ShowImage = True
        Me.ProofreadButton.ScreenTip = "对选中或全文进行语言校对"
        Me.ProofreadButton.SuperTip = "校正语法、拼写并返回可解析的修订JSON"

        ' 排版按钮
        Me.ReformatButton.Label = "排版"
        Me.ReformatButton.Name = "ReformatButton"
        Me.ReformatButton.ShowImage = True
        Me.ReformatButton.ScreenTip = "对选中或全文进行结构化排版"
        Me.ReformatButton.SuperTip = "优化标题、段落与列表并返回可解析的修订JSON"

        ' TranslateButton - 一键翻译
        Me.TranslateButton.Label = "AI翻译"
        Me.TranslateButton.Name = "TranslateButton"
        Me.TranslateButton.ShowImage = True
        Me.TranslateButton.ScreenTip = "一键翻译文档内容"
        Me.TranslateButton.SuperTip = "支持全文翻译、选区翻译、沉浸式翻译等多种模式"

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
        Me.GroupTools.Items.Add(Me.BatchDataGenButton)
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
    Protected WithEvents DoubaoButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' 在 Class BaseOfficeRibbon 的底部添加这些控件声明
    Protected WithEvents GroupTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents WebCaptureButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents SpotlightButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    Protected WithEvents BatchDataGenButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    Protected WithEvents GroupMCP As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents MCPButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' 一键翻译按钮声明
    Protected WithEvents TranslateButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' 新增：校对/排版按钮声明
    Protected WithEvents ProofreadButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ReformatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class