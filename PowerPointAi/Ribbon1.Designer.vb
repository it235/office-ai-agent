Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' ��Ӵ�����
Partial Class Ribbon1
    Inherits ShareRibbon.BaseOfficeRibbon

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms ��׫д�����֧���������
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '����������Ҫ�˵��á�
        InitializeComponent()

    End Sub

    '�����д�ͷ�����������б�
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

    '���������������
    Private components As System.ComponentModel.IContainer

    'ע��: ���¹��������������������
    '��ʹ�����������޸�����
    '��Ҫʹ�ô���༭���޸�����
    <System.Diagnostics.DebuggerStepThrough()>
    Private Overloads Sub InitializeComponent()
        Me.TabAI.Label = "PPT AI"

        ' �����ض���ͼ��
        Me.ConfigApiButton.Image = ShareRibbon.SharedResources.AiApiConfig
        Me.DataAnalysisButton.Image = ShareRibbon.SharedResources.Magic
        Me.PromptConfigButton.Image = ShareRibbon.SharedResources.Send32
        Me.ChatButton.Image = ShareRibbon.SharedResources.Chat
        Me.AboutButton.Image = ShareRibbon.SharedResources.About
        Me.ClearCacheButton.Image = ShareRibbon.SharedResources.About

        ' ���� Excel �ض�����ʾ
        Me.DataAnalysisButton.SuperTip = "��ѡ���������������ݺ�AI������������һ��sheet��"
        Me.PromptConfigButton.SuperTip = "�������ʾ�ʿ��Ը��õİ�AIȷ���Լ��Ķ�λ����������ݸ������������"
        Me.ChatButton.SuperTip = "��ʹ�ÿͻ���һ����AI�Ի���������ӱ��"

        ' ���� RibbonType
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"

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
