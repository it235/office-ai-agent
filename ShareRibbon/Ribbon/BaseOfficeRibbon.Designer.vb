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

        ' Group 1: 免费强化版 - Deepseek/Doubao
        Me.GroupDeepseek = Me.Factory.CreateRibbonGroup
        Me.DeepseekButton = Me.Factory.CreateRibbonButton()
        Me.DoubaoButton = Me.Factory.CreateRibbonButton()

        ' Group 2: 大模型配置 - 配置API/提示词/自动补全
        Me.GroupConfig = Me.Factory.CreateRibbonGroup
        Me.ConfigApiButton = Me.Factory.CreateRibbonButton
        Me.PromptConfigButton = Me.Factory.CreateRibbonButton
        Me.AutocompleteSettingsButton = Me.Factory.CreateRibbonButton

        ' Group 3: AI对话 - Chat AI/AI翻译
        Me.GroupChat = Me.Factory.CreateRibbonGroup
        Me.ChatButton = Me.Factory.CreateRibbonButton
        Me.TranslateButton = Me.Factory.CreateRibbonButton

        ' Group 4: AI内容提效 - 续写/校对/排版/模板排版/接受补全 (Word/PPT专用)
        Me.GroupAIContent = Me.Factory.CreateRibbonGroup
        Me.ContinuationButton = Me.Factory.CreateRibbonButton
        Me.ProofreadButton = Me.Factory.CreateRibbonButton
        Me.ReformatButton = Me.Factory.CreateRibbonButton
        Me.TemplateFormatButton = Me.Factory.CreateRibbonButton

        ' Group 5: MCP连接
        Me.GroupMCP = Me.Factory.CreateRibbonGroup
        Me.MCPButton = Me.Factory.CreateRibbonButton()

        ' Group 6: 关于与设置
        Me.GroupAbout = Me.Factory.CreateRibbonGroup
        Me.AboutButton = Me.Factory.CreateRibbonButton
        Me.ClearCacheButton = Me.Factory.CreateRibbonButton

        ' Group 7: 帮助与学习
        Me.GroupHelp = Me.Factory.CreateRibbonGroup
        Me.StudyButton = Me.Factory.CreateRibbonButton

        ' Group: 工具箱 (Excel专用)
        Me.GroupTools = Me.Factory.CreateRibbonGroup
        Me.DataAnalysisButton = Me.Factory.CreateRibbonButton
        Me.WebCaptureButton = Me.Factory.CreateRibbonButton
        Me.SpotlightButton = Me.Factory.CreateRibbonButton
        Me.BatchDataGenButton = Me.Factory.CreateRibbonButton()

        ' 兼容旧代码的分隔符（保留但不使用）
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        ' 兼容旧代码的GroupAI（保留但不使用）
        Me.GroupAI = Me.Factory.CreateRibbonGroup

        Me.TabAI.SuspendLayout()
        Me.SuspendLayout()

        ' ========== TabAI 布局 ==========
        Me.TabAI.Groups.Add(Me.GroupDeepseek)   ' 1. 免费强化版
        Me.TabAI.Groups.Add(Me.GroupConfig)     ' 2. 大模型配置
        Me.TabAI.Groups.Add(Me.GroupChat)       ' 3. AI对话
        Me.TabAI.Groups.Add(Me.GroupAIContent)  ' 4. AI内容提效
        Me.TabAI.Groups.Add(Me.GroupTools)      ' 5. 工具箱
        Me.TabAI.Groups.Add(Me.GroupMCP)        ' 6. MCP连接
        Me.TabAI.Groups.Add(Me.GroupAbout)      ' 7. 关于与设置
        Me.TabAI.Groups.Add(Me.GroupHelp)       ' 8. 帮助与学习

        Me.TabAI.Label = "AI助手"
        Me.TabAI.Name = "TabAI"

        ' ========== Group 1: 免费强化版 ==========
        Me.GroupDeepseek.Items.Add(Me.DeepseekButton)
        Me.GroupDeepseek.Items.Add(Me.DoubaoButton)
        Me.GroupDeepseek.Label = "免费强化版"
        Me.GroupDeepseek.Name = "GroupDeepseek"

        Me.DeepseekButton.Label = "Deepseek"
        Me.DeepseekButton.Name = "DeepseekButton"
        Me.DeepseekButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeepseekButton.ShowImage = True
        Me.DeepseekButton.ScreenTip = "免费增强版"
        Me.DeepseekButton.SuperTip = "在原有对话基础上，增加Agent执行能力"

        Me.DoubaoButton.Label = "Doubao"
        Me.DoubaoButton.Name = "DoubaoButton"
        Me.DoubaoButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DoubaoButton.ShowImage = True
        Me.DoubaoButton.ScreenTip = "豆包智能助手"
        Me.DoubaoButton.SuperTip = "基于豆包的智能对话助手，支持代码执行"

        ' ========== Group 2: 大模型配置 ==========
        Me.GroupConfig.Items.Add(Me.ConfigApiButton)
        Me.GroupConfig.Items.Add(Me.PromptConfigButton)
        Me.GroupConfig.Items.Add(Me.AutocompleteSettingsButton)
        Me.GroupConfig.Label = "大模型配置"
        Me.GroupConfig.Name = "GroupConfig"

        Me.ConfigApiButton.Label = "配置API"
        Me.ConfigApiButton.Name = "ConfigApiButton"
        Me.ConfigApiButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ConfigApiButton.ShowImage = True
        Me.ConfigApiButton.ScreenTip = "配置大模型API"
        Me.ConfigApiButton.SuperTip = "使用AI功能前需要配置apiKey"

        Me.PromptConfigButton.Label = "提示词"
        Me.PromptConfigButton.Name = "PromptConfigButton"
        Me.PromptConfigButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PromptConfigButton.ShowImage = True
        Me.PromptConfigButton.ScreenTip = "配置提示词"
        Me.PromptConfigButton.SuperTip = "管理和配置AI对话的系统提示词"

        Me.AutocompleteSettingsButton.Label = "自动补全"
        Me.AutocompleteSettingsButton.Name = "AutocompleteSettingsButton"
        Me.AutocompleteSettingsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AutocompleteSettingsButton.ShowImage = True
        Me.AutocompleteSettingsButton.ScreenTip = "配置AI自动补全"
        Me.AutocompleteSettingsButton.SuperTip = "设置自动补全开关、快捷键和触发延迟"

        ' ========== Group 3: AI对话 ==========
        Me.GroupChat.Items.Add(Me.ChatButton)
        Me.GroupChat.Items.Add(Me.TranslateButton)
        Me.GroupChat.Label = "AI对话"
        Me.GroupChat.Name = "GroupChat"

        Me.ChatButton.Label = "Chat AI"
        Me.ChatButton.Name = "ChatButton"
        Me.ChatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ChatButton.ShowImage = True
        Me.ChatButton.ScreenTip = "AI对话助手"
        Me.ChatButton.SuperTip = "打开AI对话面板，支持多轮对话和代码执行"

        Me.TranslateButton.Label = "AI翻译"
        Me.TranslateButton.Name = "TranslateButton"
        Me.TranslateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TranslateButton.ShowImage = True
        Me.TranslateButton.ScreenTip = "一键翻译文档内容"
        Me.TranslateButton.SuperTip = "支持全文翻译、选区翻译、沉浸式翻译等多种模式"

        ' ========== Group 4: AI内容提效 ==========
        Me.GroupAIContent.Items.Add(Me.ContinuationButton)
        Me.GroupAIContent.Items.Add(Me.ProofreadButton)
        Me.GroupAIContent.Items.Add(Me.ReformatButton)
        Me.GroupAIContent.Items.Add(Me.TemplateFormatButton)
        Me.GroupAIContent.Label = "AI内容提效"
        Me.GroupAIContent.Name = "GroupAIContent"

        Me.ContinuationButton.Label = "续写"
        Me.ContinuationButton.Name = "ContinuationButton"
        Me.ContinuationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ContinuationButton.ShowImage = True
        Me.ContinuationButton.ScreenTip = "AI智能续写"
        Me.ContinuationButton.SuperTip = "根据光标位置的上下文智能续写内容"

        Me.ProofreadButton.Label = "校对"
        Me.ProofreadButton.Name = "ProofreadButton"
        Me.ProofreadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ProofreadButton.ShowImage = True
        Me.ProofreadButton.ScreenTip = "对选中或全文进行语言校对"
        Me.ProofreadButton.SuperTip = "校正语法、拼写并返回可解析的修订JSON"

        Me.ReformatButton.Label = "排版"
        Me.ReformatButton.Name = "ReformatButton"
        Me.ReformatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ReformatButton.ShowImage = True
        Me.ReformatButton.ScreenTip = "对选中或全文进行结构化排版"
        Me.ReformatButton.SuperTip = "优化标题、段落与列表并返回可解析的修订JSON"

        Me.TemplateFormatButton.Label = "模板排版"
        Me.TemplateFormatButton.Name = "TemplateFormatButton"
        Me.TemplateFormatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TemplateFormatButton.ShowImage = True
        Me.TemplateFormatButton.ScreenTip = "使用模板格式排版"
        Me.TemplateFormatButton.SuperTip = "选择格式模板，AI生成内容时将参考模板中的字体、字号、段落等格式"


        ' ========== Group 5: 工具箱 (Excel专用) ==========
        Me.GroupTools.Items.Add(Me.DataAnalysisButton)
        Me.GroupTools.Items.Add(Me.WebCaptureButton)
        Me.GroupTools.Items.Add(Me.SpotlightButton)
        Me.GroupTools.Items.Add(Me.BatchDataGenButton)
        Me.GroupTools.Label = "工具箱"
        Me.GroupTools.Name = "GroupTools"

        Me.DataAnalysisButton.Label = "数据分析"
        Me.DataAnalysisButton.Name = "DataAnalysisButton"
        Me.DataAnalysisButton.ShowImage = True
        Me.DataAnalysisButton.ScreenTip = "智能数据分析"
        Me.DataAnalysisButton.SuperTip = "AI辅助分析Excel数据"

        Me.WebCaptureButton.Label = "抓取网页"
        Me.WebCaptureButton.Name = "WebCaptureButton"
        Me.WebCaptureButton.ShowImage = True
        Me.WebCaptureButton.SuperTip = "打开网页捕获工具"

        Me.SpotlightButton.Label = "聚光灯"
        Me.SpotlightButton.Name = "SpotlightButton"
        Me.SpotlightButton.ShowImage = True
        Me.SpotlightButton.SuperTip = "高亮选中单元格所在的行和列"

        Me.BatchDataGenButton.Label = "批量生成"
        Me.BatchDataGenButton.Name = "BatchDataGenButton"
        Me.BatchDataGenButton.ShowImage = True
        Me.BatchDataGenButton.ScreenTip = "配置和生成批量数据"
        Me.BatchDataGenButton.SuperTip = "配置字段、列关系并生成数据到工作簿"

        ' ========== Group 6: MCP连接 ==========
        Me.GroupMCP.Items.Add(Me.MCPButton)
        Me.GroupMCP.Label = "MCP连接"
        Me.GroupMCP.Name = "GroupMCP"

        Me.MCPButton.Label = "MCP"
        Me.MCPButton.Name = "MCPButton"
        Me.MCPButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MCPButton.ShowImage = True
        Me.MCPButton.ScreenTip = "MCP服务器配置"
        Me.MCPButton.SuperTip = "配置MCP服务器并作为客户端调用大模型"

        ' ========== Group 7: 关于与设置 ==========
        Me.GroupAbout.Items.Add(Me.AboutButton)
        Me.GroupAbout.Items.Add(Me.ClearCacheButton)
        Me.GroupAbout.Label = "关于与设置"
        Me.GroupAbout.Name = "GroupAbout"

        Me.AboutButton.Label = "关于"
        Me.AboutButton.Name = "AboutButton"
        Me.AboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AboutButton.ShowImage = True
        Me.AboutButton.ScreenTip = "关于本插件"
        Me.AboutButton.SuperTip = "查看插件信息和开源地址"

        Me.ClearCacheButton.Label = "清理缓存"
        Me.ClearCacheButton.Name = "ClearCacheButton"
        Me.ClearCacheButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ClearCacheButton.ShowImage = True
        Me.ClearCacheButton.ScreenTip = "清理配置缓存"
        Me.ClearCacheButton.SuperTip = "清除所有配置和历史记录"

        ' ========== Group 8: 帮助与学习 ==========
        Me.GroupHelp.Items.Add(Me.StudyButton)
        Me.GroupHelp.Label = "帮助与学习"
        Me.GroupHelp.Name = "GroupHelp"

        Me.StudyButton.Label = "教学文档"
        Me.StudyButton.Name = "StudyButton"
        Me.StudyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.StudyButton.ShowImage = True
        Me.StudyButton.ScreenTip = "查看教学文档"
        Me.StudyButton.SuperTip = "打开在线教学文档，了解所有功能的使用方法"

        ' ========== 兼容旧代码 ==========
        Me.Separator1.Name = "Separator1"
        Me.Separator2.Name = "Separator2"
        Me.GroupAI.Label = "AI大模型"
        Me.GroupAI.Name = "GroupAI"

        ' BaseOfficeRibbon
        Me.Name = "BaseOfficeRibbon"
        Me.Tabs.Add(Me.TabAI)

        Me.TabAI.ResumeLayout(False)
        Me.TabAI.PerformLayout()
        Me.ResumeLayout(False)
    End Sub

    ' Tab
    Protected WithEvents TabAI As Microsoft.Office.Tools.Ribbon.RibbonTab

    ' Group 1: 免费强化版
    Protected WithEvents GroupDeepseek As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents DeepseekButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents DoubaoButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 2: 大模型配置
    Protected WithEvents GroupConfig As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ConfigApiButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents PromptConfigButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents AutocompleteSettingsButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 3: AI对话
    Protected WithEvents GroupChat As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ChatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents TranslateButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 4: AI内容提效
    Protected WithEvents GroupAIContent As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ContinuationButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ProofreadButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ReformatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents TemplateFormatButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 5: 工具箱
    Protected WithEvents GroupTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents DataAnalysisButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents WebCaptureButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents SpotlightButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents BatchDataGenButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 6: MCP连接
    Protected WithEvents GroupMCP As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents MCPButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 7: 关于与设置
    Protected WithEvents GroupAbout As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents AboutButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ClearCacheButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 8: 帮助与学习
    Protected WithEvents GroupHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents StudyButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' 兼容旧代码（保留但不再使用）
    Protected WithEvents GroupAI As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Protected WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
End Class
