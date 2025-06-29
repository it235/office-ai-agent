Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.WinForms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BaseDeepseekChat
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Protected components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
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
        Me.ChatBrowser.Name = "DeepseekBrowser"
        Me.ChatBrowser.Size = New System.Drawing.Size(400, 600)
        Me.ChatBrowser.TabIndex = 1
        Me.ChatBrowser.ZoomFactor = 1.0R
        '
        '
        'DeepseekControl
        '
        Me.Controls.Add(Me.ChatBrowser)
        Me.Name = "DeepseekControl"
        Me.Size = New System.Drawing.Size(400, 600)
        CType(Me.ChatBrowser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Protected WithEvents ChatBrowser As Microsoft.Web.WebView2.WinForms.WebView2
End Class