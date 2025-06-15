<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WebDataCapturePane
    Inherits ShareRibbon.BaseDataCapturePane

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
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        ' 调用基类的初始化
        MyBase.InitializeComponent()

        ' Word 特定的初始化
        'Me.ChatBrowser.ZoomFactor = 1.25R  ' Word 可能需要更大的缩放比例
        Me.Name = "WordDataCapturePane"

        ' 可以添加 Word 特有的控件
        ' ...

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
