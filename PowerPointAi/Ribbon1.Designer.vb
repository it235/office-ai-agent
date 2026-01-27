Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Partial Class Ribbon1
    Inherits ShareRibbon.BaseOfficeRibbon

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Overloads Sub InitializeComponent()
        Me.TabAI.Label = "PPT AI"

        ' 设置特定的图标
        Me.ConfigApiButton.Image = ShareRibbon.SharedResources.AiApiConfig
        Me.DataAnalysisButton.Image = ShareRibbon.SharedResources.Magic
        Me.PromptConfigButton.Image = ShareRibbon.SharedResources.Send32
        Me.ChatButton.Image = ShareRibbon.SharedResources.Chat
        Me.AboutButton.Image = ShareRibbon.SharedResources.About
        Me.ClearCacheButton.Image = ShareRibbon.SharedResources.Clear

        ' 设置 Excel 特定的提示
        Me.DataAnalysisButton.SuperTip = "可选中提出的问题和数据后AI帮你整理到另外一个sheet中"
        Me.PromptConfigButton.SuperTip = "优秀的提示词可以更好的帮AI确定自己的定位，让输出内容更符合你的期望"
        Me.ChatButton.SuperTip = "像使用客户端一样与AI对话，聊天更加便捷"

        ' 设置 RibbonType
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"

        Me.MCPButton.Image = ShareRibbon.SharedResources.Mcp1
        Me.BatchDataGenButton.Visible = False
        Me.SpotlightButton.Visible = False
        Me.WebCaptureButton.Visible = False

        Me.DeepseekButton.Image = ShareRibbon.SharedResources.Deepseek
        Me.DoubaoButton.Image = ShareRibbon.SharedResources.Doubao
        Me.WebCaptureButton.Image = ShareRibbon.SharedResources.Send32

        Me.ContinuationButton.Visible = False
        Me.TranslateButton.Image = ShareRibbon.SharedResources.Translate
        Me.StudyButton.Image = ShareRibbon.SharedResources.Help

        Me.ProofreadButton.Visible = False
        Me.ReformatButton.Visible = False
        Me.DataAnalysisButton.Visible = False
        Me.TemplateFormatButton.Visible = False
        Me.AcceptCompletionButton.Visible = False ' 暂时不展示
    End Sub

End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
