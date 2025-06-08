Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' 订阅Word的SelectionChange 事件
        ' 帮我补全word选择的内容事件
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 PowerPoint.Selection 对象
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 根据选择类型处理内容
            If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                ' 处理形状选择
                Dim shapeRange = selection.ShapeRange
                If shapeRange.Count > 0 Then
                    ' 检查是否是表格
                    If shapeRange(1).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        ' 处理表格
                        Dim table = shapeRange(1).Table
                        Dim sb As New StringBuilder()
                        For row As Integer = 1 To table.Rows.Count
                            For col As Integer = 1 To table.Columns.Count
                                sb.Append(table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim())
                                If col < table.Columns.Count Then sb.Append(vbTab)
                            Next
                            sb.AppendLine()
                        Next
                        content = sb.ToString()
                    Else
                        ' 处理普通形状
                        content = "[已选中 " & shapeRange.Count & " 个形状]"
                        For i = 1 To shapeRange.Count
                            If shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                content &= vbCrLf & shapeRange(i).TextFrame.TextRange.Text
                            End If
                        Next
                    End If
                End If

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                ' 处理文本选择
                content = selection.TextRange.Text

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                ' 处理幻灯片选择
                content = "[已选中 " & selection.SlideRange.Count & " 张幻灯片]"
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "PowerPoint幻灯片",  ' 使用文档名称作为标识
                content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", "")
            )
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取PowerPoint选中内容时出错: {ex.Message}")
        End Try
    End Sub

    Private Function GetSelectionDetails(selection As Object) As String
        Try
            Dim details As New StringBuilder()
            Dim ppSelection = TryCast(selection, Microsoft.Office.Interop.PowerPoint.Selection)

            If ppSelection Is Nothing Then
                Return "未选中任何内容"
            End If

            ' 添加基本信息
            details.AppendLine($"选择类型: {ppSelection.Type}")

            If ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                Dim shapeRange = ppSelection.ShapeRange
                details.AppendLine($"形状数量: {shapeRange.Count}")
                For i = 1 To shapeRange.Count
                    details.AppendLine($"形状 {i} 类型: {shapeRange(i).Type}")
                    ' 检查是否是表格
                    If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim table = shapeRange(i).Table
                        details.AppendLine($"表格大小: {table.Rows.Count}行 x {table.Columns.Count}列")
                    ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        details.AppendLine($"形状 {i} 文本长度: {shapeRange(i).TextFrame.TextRange.Length}")
                    End If
                Next

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                Dim textRange = ppSelection.TextRange
                details.AppendLine($"文本长度: {textRange.Length}")
                details.AppendLine($"字符数: {textRange.Length}")

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                Dim slideRange = ppSelection.SlideRange
                details.AppendLine($"选中幻灯片数: {slideRange.Count}")
                For i = 1 To slideRange.Count
                    details.AppendLine($"幻灯片 {i} 标题: {slideRange(i).Name}")
                Next
            End If

            Return details.ToString()
        Catch ex As Exception
            Return $"获取选择详情时出错: {ex.Message}"
        End Try
    End Function

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


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

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' 这里可以实现word的特殊逻辑
        Send(message)
    End Sub


End Class

