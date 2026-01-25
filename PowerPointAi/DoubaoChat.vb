' PowerPointAi\DoubaoChat.vb
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Tools
Imports Microsoft.Vbe.Interop
Imports ShareRibbon

Public Class DoubaoChat
    Inherits BaseDoubaoChat

    Public Sub New()
        ' 这将调用InitializeComponent方法
        InitializeComponent()
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        'InitializeWebView2Script()
    End Sub
    'Public Async Function InitializeAsync() As Task
    '    Await InitializeWebView2()
    'End Function

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            If pptApp.ActivePresentation IsNot Nothing Then
                Dim path As String = pptApp.ActivePresentation.Path
                If Not String.IsNullOrEmpty(path) Then
                    Return path
                End If
            End If
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Catch ex As Exception
            Debug.WriteLine("获取当前工作目录失败: " & ex.Message)
            Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Try
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            Dim selection As Object = pptApp.ActiveWindow.Selection
            
            If selection IsNot Nothing Then
                Dim sb As New StringBuilder()
                
                ' 获取选中的内容
                If selection.Type = PpSelectionType.ppSelectionText Then
                    ' 文本选择
                    Dim textRange As TextRange = selection.TextRange
                    If textRange IsNot Nothing Then
                        Dim selectedText As String = textRange.Text.Trim()
                        If Not String.IsNullOrEmpty(selectedText) Then
                            Return message & vbCrLf & vbCrLf & "当前选中的文本:" & vbCrLf & selectedText
                        End If
                    End If
                ElseIf selection.Type = PpSelectionType.ppSelectionShapes Then
                    ' 形状选择
                    Dim shapes As ShapeRange = selection.ShapeRange
                    If shapes IsNot Nothing Then
                        sb.AppendLine("选中的形状:")
                        For i As Integer = 1 To shapes.Count
                            Dim shape As Shape = shapes(i)
                            sb.AppendLine($"形状 {i}: {shape.Name}")
                            
                            ' 获取形状中的文本
                            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                               shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                Dim shapeText As String = shape.TextFrame.TextRange.Text.Trim()
                                If Not String.IsNullOrEmpty(shapeText) Then
                                    sb.AppendLine($"  文本: {shapeText}")
                                End If
                            End If
                        Next
                        
                        Dim selectedInfo As String = sb.ToString().Trim()
                        If Not String.IsNullOrEmpty(selectedInfo) Then
                            Return message & vbCrLf & vbCrLf & selectedInfo
                        End If
                    End If
                ElseIf selection.Type = PpSelectionType.ppSelectionSlides Then
                    ' 幻灯片选择
                    Dim slideRange As SlideRange = selection.SlideRange
                    If slideRange IsNot Nothing Then
                        sb.AppendLine("选中的幻灯片:")
                        For i As Integer = 1 To slideRange.Count
                            Dim slide As Slide = slideRange(i)
                            sb.AppendLine($"幻灯片 {slide.SlideIndex}: {slide.Name}")
                        Next
                        
                        Dim selectedInfo As String = sb.ToString().Trim()
                        If Not String.IsNullOrEmpty(selectedInfo) Then
                            Return message & vbCrLf & vbCrLf & selectedInfo
                        End If
                    End If
                End If
            End If
            
            Return message
        Catch ex As Exception
            Debug.WriteLine("获取选中内容失败: " & ex.Message)
            Return message
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    Protected Overrides Function GetVBProject() As Microsoft.Vbe.Interop.VBProject
        Try
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            If pptApp.ActivePresentation IsNot Nothing Then
                Return TryCast(pptApp.ActivePresentation.VBProject, Microsoft.Vbe.Interop.VBProject)
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("获取VB项目失败: " & ex.Message)
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

    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        Return True
    End Function

    'Protected Overrides Function RunCode(vbaCode As String) As Object
    '    Try
    '        Dim vbProject As VBProject = GetVBProject()
    '        If vbProject IsNot Nothing Then
    '            ' 创建新的模块来执行代码
    '            Dim moduleToAdd As VBComponent = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
    '            moduleToAdd.Name = "TempDoubaoModule_" & DateTime.Now.Ticks

    '            ' 添加代码到模块
    '            moduleToAdd.CodeModule.AddFromString(vbaCode)

    '            ' 这里可以执行代码，但需要更复杂的逻辑来调用过程
    '            ' 简化起见，我们只是记录代码已添加
    '            Debug.WriteLine("VBA代码已添加到临时模块: " & moduleToAdd.Name)

    '            ' 可以选择清理临时模块
    '            ' vbProject.VBComponents.Remove(moduleToAdd)

    '            Return True
    '        Else
    '            GlobalStatusStrip.ShowWarning("无法获取VB项目，请确保已启用宏")
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        GlobalStatusStrip.ShowWarning("执行VBA代码失败: " & ex.Message)
    '        Return False
    '    End Try
    'End Function

    Protected Overrides Sub SendChatMessage(message As String)
        Try
            ' 这里可以实现发送消息到Doubao的逻辑
            ' 但由于使用WebView2，这个方法可能不需要
            Debug.WriteLine("发送消息: " & message)
        Catch ex As Exception
            Debug.WriteLine("发送消息失败: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            ' 这里可以实现获取选择内容的逻辑
            Debug.WriteLine("获取选择内容")
        Catch ex As Exception
            Debug.WriteLine("获取选择内容失败: " & ex.Message)
        End Try
    End Sub

    ' 重写GetOfficeApplicationObject方法
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Try
            Return Globals.ThisAddIn.Application
        Catch ex As Exception
            Debug.WriteLine("获取PowerPoint应用对象失败: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class