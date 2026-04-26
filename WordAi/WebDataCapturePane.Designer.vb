<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WebDataCapturePane
    Inherits ShareRibbon.BaseDataCapturePane

    'UserControl ��д Dispose������������б���
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
    Private Shadows components As System.ComponentModel.IContainer

    'ע��: ���¹����� Windows ����������������
    '����ʹ�� Windows ����������޸�����
    '��Ҫʹ�ô���༭���޸�����
    <System.Diagnostics.DebuggerStepThrough()>
    Private Overloads Sub InitializeComponent()
        ' ���û���ĳ�ʼ��
        MyBase.InitializeComponent()

        ' Word �ض��ĳ�ʼ��
        'Me.ChatBrowser.ZoomFactor = 1.25R  ' Word ������Ҫ��������ű���
        Me.Name = "WordDataCapturePane"

        ' �������� Word ���еĿؼ�
        ' ...

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
