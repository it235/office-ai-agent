Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.WinForms
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BaseDataCapturePane
    Inherits System.Windows.Forms.UserControl

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                If components IsNot Nothing Then
                    components.Dispose()
                End If
                If ChatBrowser IsNot Nothing Then
                    ChatBrowser.Dispose()
                End If
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Protected components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub InitializeComponent()
        Me.ChatBrowser = New Microsoft.Web.WebView2.WinForms.WebView2()
        Me.TopPanel = New System.Windows.Forms.Panel()
        Me.ButtonPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.NavigateButton = New System.Windows.Forms.Button()
        Me.CaptureButton = New System.Windows.Forms.Button()
        Me.SelectDomButton = New System.Windows.Forms.Button()
        Me.UrlTextBox = New System.Windows.Forms.TextBox()
        CType(Me.ChatBrowser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TopPanel.SuspendLayout()
        Me.ButtonPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'ChatBrowser
        '
        Me.ChatBrowser.AllowExternalDrop = True
        Me.ChatBrowser.CreationProperties = Nothing
        Me.ChatBrowser.DefaultBackgroundColor = System.Drawing.Color.White
        Me.ChatBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ChatBrowser.Location = New System.Drawing.Point(0, 30)
        Me.ChatBrowser.Name = "ChatBrowser"
        Me.ChatBrowser.Size = New System.Drawing.Size(400, 570)
        Me.ChatBrowser.TabIndex = 0
        Me.ChatBrowser.ZoomFactor = 1.0R
        '
        'TopPanel
        '
        Me.TopPanel.BackColor = System.Drawing.SystemColors.Control
        Me.TopPanel.Controls.Add(Me.ButtonPanel)
        Me.TopPanel.Controls.Add(Me.UrlTextBox)
        Me.TopPanel.Dock = System.Windows.Forms.DockStyle.Top
        Me.TopPanel.Location = New System.Drawing.Point(0, 0)
        Me.TopPanel.Name = "TopPanel"
        Me.TopPanel.Padding = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.TopPanel.Size = New System.Drawing.Size(400, 30)
        Me.TopPanel.TabIndex = 1
        '
        'ButtonPanel
        '
        Me.ButtonPanel.Controls.Add(Me.NavigateButton)
        Me.ButtonPanel.Controls.Add(Me.CaptureButton)
        Me.ButtonPanel.Controls.Add(Me.SelectDomButton)
        Me.ButtonPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ButtonPanel.Location = New System.Drawing.Point(154, 3)
        Me.ButtonPanel.Name = "ButtonPanel"
        Me.ButtonPanel.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.ButtonPanel.Size = New System.Drawing.Size(241, 24)
        Me.ButtonPanel.TabIndex = 0
        Me.ButtonPanel.WrapContents = False
        '
        'NavigateButton
        '
        Me.NavigateButton.AutoSize = True
        Me.NavigateButton.Location = New System.Drawing.Point(5, 0)
        Me.NavigateButton.Margin = New System.Windows.Forms.Padding(0, 0, 5, 0)
        Me.NavigateButton.Name = "NavigateButton"
        Me.NavigateButton.Size = New System.Drawing.Size(75, 24)
        Me.NavigateButton.TabIndex = 0
        Me.NavigateButton.Text = "打开"
        '
        'CaptureButton
        '
        Me.CaptureButton.AutoSize = True
        Me.CaptureButton.Location = New System.Drawing.Point(85, 0)
        Me.CaptureButton.Margin = New System.Windows.Forms.Padding(0, 0, 5, 0)
        Me.CaptureButton.Name = "CaptureButton"
        Me.CaptureButton.Size = New System.Drawing.Size(75, 24)
        Me.CaptureButton.TabIndex = 1
        Me.CaptureButton.Text = "抓取整页"
        '
        'SelectDomButton
        '
        Me.SelectDomButton.AutoSize = True
        Me.SelectDomButton.Location = New System.Drawing.Point(165, 0)
        Me.SelectDomButton.Margin = New System.Windows.Forms.Padding(0)
        Me.SelectDomButton.Name = "SelectDomButton"
        Me.SelectDomButton.Size = New System.Drawing.Size(75, 24)
        Me.SelectDomButton.TabIndex = 2
        Me.SelectDomButton.Text = "选择元素"
        '
        'UrlTextBox
        '
        Me.UrlTextBox.Dock = System.Windows.Forms.DockStyle.Left
        Me.UrlTextBox.Location = New System.Drawing.Point(5, 3)
        Me.UrlTextBox.Name = "UrlTextBox"
        Me.UrlTextBox.Size = New System.Drawing.Size(149, 21)
        Me.UrlTextBox.TabIndex = 1
        '
        'BaseDataCapturePane
        '
        Me.Controls.Add(Me.ChatBrowser)
        Me.Controls.Add(Me.TopPanel)
        Me.Name = "BaseDataCapturePane"
        Me.Size = New System.Drawing.Size(400, 600)
        CType(Me.ChatBrowser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TopPanel.ResumeLayout(False)
        Me.TopPanel.PerformLayout()
        Me.ButtonPanel.ResumeLayout(False)
        Me.ButtonPanel.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Protected WithEvents ChatBrowser As Microsoft.Web.WebView2.WinForms.WebView2
    Protected WithEvents TopPanel As Panel
    Protected WithEvents ButtonPanel As FlowLayoutPanel
    Protected WithEvents UrlTextBox As TextBox
    Protected WithEvents NavigateButton As Button
    Protected WithEvents CaptureButton As Button
    Protected WithEvents SelectDomButton As Button
End Class