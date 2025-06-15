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

            ' 转换为 Word.Selection 对象
            Dim selection = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Word.Selection)
            If selection Is Nothing Then
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 检查是否选中了表格
            If selection.Tables.Count > 0 Then
                ' 如果选中的是表格
                Dim table = selection.Tables(1)
                Dim sb As New StringBuilder()

                ' 遍历表格内容
                For row As Integer = 1 To table.Rows.Count
                    For col As Integer = 1 To table.Columns.Count
                        sb.Append(table.Cell(row, col).Range.Text.TrimEnd(ChrW(13), ChrW(7)))
                        If col < table.Columns.Count Then sb.Append(vbTab)
                    Next
                    sb.AppendLine()
                Next
                content = sb.ToString()

            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' 如果选中的是图片或形状
                content = "[图片或形状]"
            Else
                ' 普通文本选择
                content = selection.Text
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "Word文档",  ' 使用文档名称作为标识
                If(selection.Tables.Count > 0,
                   "[表格内容]",
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
            )
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取Word选中内容时出错: {ex.Message}")
        End Try
    End Sub


    ' 获取选中内容的详细信息
    Private Function GetSelectionDetails(selection As Microsoft.Office.Interop.Word.Selection) As String
        Dim details As New StringBuilder()

        ' 添加基本信息
        details.AppendLine($"开始位置: {selection.Start}")
        details.AppendLine($"结束位置: {selection.End}")
        details.AppendLine($"字符数: {selection.Characters.Count}")

        ' 如果是表格，添加表格信息
        If selection.Tables.Count > 0 Then
            Dim table = selection.Tables(1)
            details.AppendLine($"表格大小: {table.Rows.Count}行 x {table.Columns.Count}列")
        End If

        Return details.ToString()
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
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' 这里可以实现word的特殊逻辑
        Send(message)
    End Sub


    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 创建一个 Word 应用程序实例
            Dim wordApp As New Microsoft.Office.Interop.Word.Application
            wordApp.Visible = False

            Dim document As Microsoft.Office.Interop.Word.Document = Nothing
            Try
                document = wordApp.Documents.Open(filePath, ReadOnly:=True)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"文件: {Path.GetFileName(filePath)} 包含以下内容:")

                ' 获取文档文本
                Dim text As String = document.Content.Text

                ' 限制文本长度
                Dim maxTextLength As Integer = 2000
                If text.Length > maxTextLength Then
                    contentBuilder.AppendLine(text.Substring(0, maxTextLength) & "...")
                    contentBuilder.AppendLine($"[文档太长，只显示前 {maxTextLength} 个字符，总长度: {text.Length} 个字符]")
                Else
                    contentBuilder.AppendLine(text)
                End If

                Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = contentBuilder.ToString(),
                .RawData = Nothing
            }

            Finally
                ' 清理资源
                If document IsNot Nothing Then
                    document.Close(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(document)
                End If

                wordApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析 Word 文件时出错: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Word",
            .ParsedContent = $"[解析 Word 文件时出错: {ex.Message}]"
        }
        End Try
    End Function
    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            ' 获取当前活动工作簿的路径
            If Globals.ThisAddIn.Application.ActiveWorkbook IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取当前工作目录时出错: {ex.Message}")
        End Try

        ' 如果无法获取工作簿路径，则返回应用程序目录
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' 检查是否启用了选择功能
            If Not selectedCellChecked Then
                Return message
            End If

            ' 获取当前 Word 文档中的选择
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection Is Nothing Then
                Return message
            End If

            ' 创建内容构建器，格式化选中内容
            Dim contentBuilder As New StringBuilder()
            contentBuilder.AppendLine(vbCrLf & "--- 用户选中的 Word 内容 ---")

            ' 添加文档信息
            Dim activeDocument = Globals.ThisAddIn.Application.ActiveDocument
            If activeDocument IsNot Nothing Then
                contentBuilder.AppendLine($"文档: {Path.GetFileName(activeDocument.FullName)}")
            End If

            ' 选择范围信息
            contentBuilder.AppendLine($"选择范围: 第 {selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterLineNumber)} 行起")
            contentBuilder.AppendLine($"选中字符数: {selection.Characters.Count}")

            ' 处理选中内容
            If selection.Tables.Count > 0 Then
                ' 处理表格
                contentBuilder.AppendLine("选中内容类型: 表格")
                AppendTableContent(contentBuilder, selection)
            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' 处理图片或形状
                contentBuilder.AppendLine("选中内容类型: 图片或形状")
                contentBuilder.AppendLine("[图片或形状内容无法直接转换为文本]")
            Else
                ' 处理普通文本
                contentBuilder.AppendLine("选中内容类型: 文本")
                Dim text As String = selection.Text.Trim()

                ' 限制文本长度
                Dim maxLength As Integer = 2000
                If text.Length > maxLength Then
                    contentBuilder.AppendLine(text.Substring(0, maxLength) & "...")
                    contentBuilder.AppendLine($"[选中文本太长，只显示前 {maxLength} 个字符，总长度: {text.Length} 个字符]")
                Else
                    contentBuilder.AppendLine(text)
                End If
            End If

            contentBuilder.AppendLine("--- 选中内容结束 ---" & vbCrLf)

            ' 返回原始消息加上选中内容
            Return message & contentBuilder.ToString()

        Catch ex As Exception
            Debug.WriteLine($"处理Word选中内容时出错: {ex.Message}")
            Return message ' 出错时返回原始消息
        End Try
    End Function

    ' 辅助方法：处理表格内容
    Private Sub AppendTableContent(builder As StringBuilder, selection As Microsoft.Office.Interop.Word.Selection)
        Try
            ' 获取选中的表格
            Dim table As Microsoft.Office.Interop.Word.Table = Nothing

            ' 表格可能有两种情况：1. 选中了整个表格 2. 选中了表格中的单元格
            If selection.Tables.Count > 0 Then
                table = selection.Tables(1)
            ElseIf selection.Cells.Count > 0 Then
                ' 如果只选中了单元格，获取包含这些单元格的表格
                table = selection.Cells(1).Range.Tables(1)
            End If

            If table Is Nothing Then
                builder.AppendLine("[无法获取表格内容]")
                Return
            End If

            ' 添加表格信息
            builder.AppendLine($"表格大小: {table.Rows.Count} 行 × {table.Columns.Count} 列")
            builder.AppendLine()

            ' 限制显示的行列数
            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

            ' 处理表格头部（表格第一行）
            If table.Rows.Count > 0 Then
                ' 构建表头分隔线
                Dim headerBuilder As New StringBuilder()
                Dim separatorBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(1, col).Range.Text
                        ' 移除特殊字符
                        cellText = cellText.TrimEnd(ChrW(13), ChrW(7), ChrW(9), ChrW(10), ChrW(32))

                        ' 限制单元格文本长度
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' 填充表头
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append(cellText)
                        separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                    Catch ex As Exception
                        ' 忽略单元格处理错误
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append("N/A")
                        separatorBuilder.Append("---")
                    End Try
                Next

                ' 添加表头和分隔线
                builder.AppendLine(headerBuilder.ToString())
                builder.AppendLine(separatorBuilder.ToString())
            End If

            ' 处理表格数据行
            For row As Integer = 2 To maxRows ' 从第2行开始（跳过表头）
                Dim rowBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(row, col).Range.Text
                        ' 移除特殊字符
                        cellText = cellText.TrimEnd(ChrW(13), ChrW(7), ChrW(9), ChrW(10), ChrW(32))

                        ' 限制单元格文本长度
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' 填充行数据
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append(cellText)
                    Catch ex As Exception
                        ' 忽略单元格处理错误
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append("N/A")
                    End Try
                Next

                ' 添加行数据
                builder.AppendLine(rowBuilder.ToString())
            Next

            ' 如果有更多行未显示，添加提示
            If table.Rows.Count > maxRows Then
                builder.AppendLine($"... [表格共有 {table.Rows.Count} 行，仅显示前 {maxRows} 行]")
            End If

            ' 如果有更多列未显示，添加提示
            If table.Columns.Count > maxCols Then
                builder.AppendLine($"... [表格共有 {table.Columns.Count} 列，仅显示前 {maxCols} 列]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[处理表格内容时出错: {ex.Message}]")
        End Try
    End Sub
End Class

