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


    ' 实现Excel公式评估' 执行Excel公式或函数 - 增强版支持赋值和预览
    Protected Overrides Function EvaluateFormula(formulaCode As String, preview As Boolean) As Boolean
        Try
            ' 检查是否是赋值语句 (例如 C1=A1+B1)
            Dim isAssignment As Boolean = Regex.IsMatch(formulaCode, "^[A-Za-z]+[0-9]+\s*=")

            If isAssignment Then
                ' 解析赋值语句
                Dim parts As String() = formulaCode.Split(New Char() {"="c}, 2)
                Dim targetCell As String = parts(0).Trim()
                Dim formula As String = parts(1).Trim()

                ' 如果公式以=开头，则移除
                If formula.StartsWith("=") Then
                    formula = formula.Substring(1)
                End If

                ' 如果需要预览，显示预览对话框
                If preview Then
                    Dim excel As Object = Globals.ThisAddIn.Application
                    Dim currentValue As Object = Nothing
                    Try
                        currentValue = excel.Range(targetCell).Value
                    Catch ex As Exception
                        ' 单元格可能不存在值
                    End Try

                    ' 计算新值
                    Dim newValue As Object = excel.Evaluate(formula)

                    ' 创建预览对话框
                    Dim previewMsg As String = $"将要在单元格 {targetCell} 中应用公式:" & vbCrLf & vbCrLf &
                                          $"={formula}" & vbCrLf & vbCrLf &
                                          $"当前值: {If(currentValue Is Nothing, "(空)", currentValue)}" & vbCrLf &
                                          $"新值: {If(newValue Is Nothing, "(空)", newValue)}"

                    Dim result As DialogResult = MessageBox.Show(previewMsg, "Excel公式预览",
                                                          MessageBoxButtons.OKCancel,
                                                          MessageBoxIcon.Information)

                    If result <> DialogResult.OK Then
                        Return False
                    End If
                End If

                ' 执行赋值
                Dim range As Object = Globals.ThisAddIn.Application.Range(targetCell)
                range.Formula = "=" & formula

                GlobalStatusStrip.ShowInfo($"公式 '={formula}' 已应用到单元格 {targetCell}")
                Return True
            Else
                ' 普通公式计算 (不包含赋值)
                ' 去除可能的等号前缀
                If formulaCode.StartsWith("=") Then
                    formulaCode = formulaCode.Substring(1)
                End If

                ' 计算公式结果
                Dim result As Object = Globals.ThisAddIn.Application.Evaluate(formulaCode)

                ' 如果需要预览，显示计算结果
                If preview Then
                    Dim previewMsg As String = $"公式计算结果:" & vbCrLf & vbCrLf &
                                         $"={formulaCode}" & vbCrLf & vbCrLf &
                                         $"结果: {If(result Is Nothing, "(空)", result)}"

                    MessageBox.Show(previewMsg, "Excel公式结果", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' 显示结果
                    GlobalStatusStrip.ShowInfo($"公式 '={formulaCode}' 的计算结果: {result}")
                End If

                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("执行Excel公式时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
End Class
