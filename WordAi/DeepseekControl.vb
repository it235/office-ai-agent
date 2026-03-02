Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop
Imports ShareRibbon

Public Class DeepseekControl
    Inherits BaseDeepseekChat


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        'InitializeWebView2Script()
    End Sub

    Protected Overrides Sub SendChatMessage(message As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Sub GetSelectionContent(target As Object)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Throw New NotImplementedException()
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Throw New NotImplementedException()
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Throw New NotImplementedException()
    End Function

    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    Protected Overrides Function RunCode(code As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(code)
            Return True
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return False
        Catch ex As Exception
            MessageBox.Show("执行代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function


    ' 执行前预览代码
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        Return True
    End Function

    ' 提供Excel应用程序对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function


    ' Word 不支持公式评估，此功能仅适用于 Excel
    Protected Overrides Function EvaluateFormula(formulaCode As String, preview As Boolean) As Boolean
        GlobalStatusStrip.ShowWarning("公式评估功能不支持在 Word 中使用")
        Return False
    End Function
End Class
