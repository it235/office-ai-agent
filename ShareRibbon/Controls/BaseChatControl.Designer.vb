' ShareRibbon\Controls\BaseChatControl.Designer.vb
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.WinForms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BaseChatControl
    Inherits System.Windows.Forms.UserControl

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
    Protected components As System.ComponentModel.IContainer

    'ע��: ���¹����� Windows ����������������
    '����ʹ�� Windows ����������޸�����
    '��Ҫʹ�ô���༭���޸�����
    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub InitializeComponent()
        Me.ChatBrowser = New Microsoft.Web.WebView2.WinForms.WebView2()
        CType(Me.ChatBrowser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ChatBrowser
        '
        Me.ChatBrowser.AllowExternalDrop = True
        Me.ChatBrowser.CreationProperties = Nothing
        Me.ChatBrowser.DefaultBackgroundColor = System.Drawing.Color.White
        Me.ChatBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ChatBrowser.Location = New System.Drawing.Point(0, 0)
        Me.ChatBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.ChatBrowser.Name = "ChatBrowser"
        Me.ChatBrowser.Size = New System.Drawing.Size(400, 600)
        Me.ChatBrowser.TabIndex = 1
        Me.ChatBrowser.ZoomFactor = 1.0R
        '
        'SelectedContentFlowPanel
        '
        '
        'BaseChatControl
        '
        Me.Controls.Add(Me.ChatBrowser)
        Me.Name = "BaseChatControl"
        Me.Size = New System.Drawing.Size(400, 600)
        CType(Me.ChatBrowser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Protected WithEvents ChatBrowser As Microsoft.Web.WebView2.WinForms.WebView2
End Class