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

        ' �ڶ���Group�Ͱ�ť
        Me.GroupTools = Me.Factory.CreateRibbonGroup
        Me.WebCaptureButton = Me.Factory.CreateRibbonButton
        Me.SpotlightButton = Me.Factory.CreateRibbonButton


        Me.TabAI.SuspendLayout()
        Me.GroupAI.SuspendLayout()
        Me.SuspendLayout()

        ' TabAI
        Me.TabAI.Groups.Add(Me.GroupAI)
        Me.TabAI.Groups.Add(Me.GroupTools)

        Me.TabAI.Label = "AI����"
        Me.TabAI.Name = "TabAI"

        ' GroupAI
        Me.GroupAI.Items.Add(Me.ConfigApiButton)
        Me.GroupAI.Items.Add(Me.DataAnalysisButton)
        Me.GroupAI.Items.Add(Me.Separator1)
        Me.GroupAI.Items.Add(Me.PromptConfigButton)
        Me.GroupAI.Items.Add(Me.ChatButton)
        Me.GroupAI.Items.Add(Me.Separator2)
        Me.GroupAI.Items.Add(Me.AboutButton)
        Me.GroupAI.Items.Add(Me.ClearCacheButton)
        Me.GroupAI.Label = "AI����"
        Me.GroupAI.Name = "GroupAI"

        ' ConfigApiButton
        Me.ConfigApiButton.Label = "����API"
        Me.ConfigApiButton.Name = "ConfigApiButton"
        Me.ConfigApiButton.ShowImage = True
        Me.ConfigApiButton.SuperTip = "ʹ��AI����ǰ��Ҫ����apiKey"

        ' DataAnalysisButton
        Me.DataAnalysisButton.Label = "���ݷ���"
        Me.DataAnalysisButton.Name = "DataAnalysisButton"
        Me.DataAnalysisButton.ShowImage = True

        ' PromptConfigButton
        Me.PromptConfigButton.Label = "��ʾ��"
        Me.PromptConfigButton.Name = "PromptConfigButton"
        Me.PromptConfigButton.ShowImage = True

        ' ChatButton
        Me.ChatButton.Label = "AI����"
        Me.ChatButton.Name = "ChatButton"
        Me.ChatButton.ShowImage = True

        ' AboutButton
        Me.AboutButton.Label = "����"
        Me.AboutButton.Name = "AboutButton"
        Me.AboutButton.ShowImage = True

        ' ClearCacheButton
        Me.ClearCacheButton.Label = "������"
        Me.ClearCacheButton.Name = "ClearCacheButton"
        Me.ClearCacheButton.ShowImage = True

        ' Separators
        Me.Separator1.Name = "Separator1"
        Me.Separator2.Name = "Separator2"

        ' �ڶ���Group
        ' GroupTools
        Me.GroupTools.Items.Add(Me.WebCaptureButton)
        Me.GroupTools.Items.Add(Me.SpotlightButton)
        Me.GroupTools.Label = "������"
        Me.GroupTools.Name = "GroupTools"

        ' WebCaptureButton
        Me.WebCaptureButton.Label = "ץȡ��ҳ"
        Me.WebCaptureButton.Name = "WebCaptureButton"
        Me.WebCaptureButton.ShowImage = True
        Me.WebCaptureButton.SuperTip = "����ҳ���񹤾�"

        ' SpotlightButton
        Me.SpotlightButton.Label = "�۹��"
        Me.SpotlightButton.Name = "SpotlightButton"
        Me.SpotlightButton.ShowImage = True
        Me.SpotlightButton.SuperTip = "����ѡ�е�Ԫ�����ڵ��к���"


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

    ' �� Class BaseOfficeRibbon �ĵײ������Щ�ؼ�����
    Protected WithEvents GroupTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents WebCaptureButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents SpotlightButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class