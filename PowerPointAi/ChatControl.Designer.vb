' WordAi\ChatControl.Designer.vb
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ChatControl
    Inherits ShareRibbon.BaseChatControl


    'UserControl ��д Dispose������������б�
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

    'Windows ����������������
    Private components As System.ComponentModel.IContainer

    'ע��: ���¹����� Windows ����������������
    '����ʹ�� Windows ����������޸�����
    '��Ҫʹ�ô���༭���޸�����
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        ' ���û���ĳ�ʼ��
        MyBase.InitializeComponent()

        ' Word �ض��ĳ�ʼ��
        'Me.ChatBrowser.ZoomFactor = 1.25R  ' Word ������Ҫ��������ű���
        Me.Name = "WordChatControl"

        ' ������� Word ���еĿؼ�
        ' ...

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class