' ExcelAi\DoubaoChat.vb
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
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
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
            If excelApp.ActiveWorkbook IsNot Nothing Then
                Dim path As String = excelApp.ActiveWorkbook.Path
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
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
            Dim selection As Microsoft.Office.Interop.Excel.Range = TryCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range)

            If selection IsNot Nothing Then
                Dim sb As New StringBuilder()

                ' 获取选中的数据
                For row As Integer = 1 To selection.Rows.Count
                    Dim rowData As New StringBuilder()
                    For col As Integer = 1 To selection.Columns.Count
                        Dim cellValue As Object = CType(selection.Cells(row, col), Microsoft.Office.Interop.Excel.Range).Value
                        If cellValue IsNot Nothing Then
                            rowData.Append(cellValue.ToString())
                        End If
                        If col < selection.Columns.Count Then
                            rowData.Append(vbTab)
                        End If
                    Next
                    sb.AppendLine(rowData.ToString())
                Next

                Dim selectedText As String = sb.ToString().Trim()
                If Not String.IsNullOrEmpty(selectedText) Then
                    Return message & vbCrLf & vbCrLf & "当前选中的数据:" & vbCrLf & selectedText
                End If
            End If

            Return message
        Catch ex As Exception
            Debug.WriteLine("获取选中内容失败: " & ex.Message)
            Return message
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function

    Protected Overrides Function GetVBProject() As Microsoft.Vbe.Interop.VBProject
        Try
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
            If excelApp.ActiveWorkbook IsNot Nothing Then
                Return TryCast(excelApp.ActiveWorkbook.VBProject, Microsoft.Vbe.Interop.VBProject)
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("获取VB项目失败: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        ' 如果需要预览
        Dim previewTool As New EnhancedPreviewAndConfirm()
        ' 允许用户预览代码变更
        If previewTool.PreviewAndConfirmVbaExecution(vbaCode) Then
            Debug.Print("预览结束，用户同意执行代码: " & vbaCode)
            Return True
        Else
            ' 用户取消或拒绝
            Return False
        End If
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
            Debug.WriteLine("获取Excel应用对象失败: " & ex.Message)
            Return Nothing
        End Try
    End Function
End Class