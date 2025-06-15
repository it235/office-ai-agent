Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Vbe.Interop
Imports ShareRibbon
Public Class WebDataCapturePane
    Inherits BaseDataCapturePane

    Private isViewInitialized As Boolean = False
    Public Sub New()
        MyBase.New()
        ' 创建 ChatControl 实例
        ' 订阅AI聊天请求事件
        AddHandler AiChatRequested, AddressOf HandleAiChatRequest
        ' 直接调用异步初始化方法
        InitializeWebViewAsync()
    End Sub

    ' 新增：异步初始化方法
    ' 异步初始化方法
    Private Async Sub InitializeWebViewAsync()
        Try
            Debug.WriteLine("Starting WebView initialization from WebDataCapturePane")
            ' 调用基类的初始化方法
            Await InitializeWebView2()
        Catch ex As Exception
            MessageBox.Show($"初始化网页视图失败: {ex.Message}", "错误",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub HandleAiChatRequest(sender As Object, content As String)
        ' 显示聊天窗口
        Globals.ThisAddIn.ShowChatTaskPane()
        ' 添加选中的内容到引用区
        Globals.ThisAddIn.chatControl.AddSelectedContentItem(
                "来自网页",  ' 使用文档名称作为标识
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
    End Sub

    ' 处理表格创建
    Protected Overrides Function CreateTable(tableData As TableData) As String
        Try
            ' 获取当前文档和选定范围
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = doc.Application.Selection

            ' 创建表格
            Dim table = doc.Tables.Add(
                Range:=selection.Range,
                NumRows:=tableData.Rows,
                NumColumns:=tableData.Columns)

            ' 填充数据
            For i = 0 To tableData.Data.Count - 1
                For j = 0 To tableData.Data(i).Count - 1
                    table.Cell(i + 1, j + 1).Range.Text = tableData.Data(i)(j)
                Next
            Next

            ' 如果有表头，设置表头样式
            If tableData.Headers.Count > 0 Then
                table.Rows(1).HeadingFormat = True
                table.Rows(1).Range.Bold = True
            End If

            ' 设置表格样式
            table.Style = "网格型"
            table.AllowAutoFit = True
            table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

            Return "[表格已插入]" & vbCrLf
        Catch ex As Exception
            MessageBox.Show($"创建表格时出错: {ex.Message}", "错误")
            Return String.Empty
        End Try
    End Function

    Protected Overrides Sub HandleExtractedContent(content As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                ' 在当前光标位置插入内容
                Dim selection = doc.Application.Selection
                If selection IsNot Nothing Then
                    ' 插入内容
                    selection.TypeText(content)
                    'selection.TypeText(vbCrLf & vbCrLf)

                    ' 插入分隔线
                    'selection.TypeText(vbCrLf & "----------------------------------------" & vbCrLf)
                    'selection.TypeText("来源: " & ChatBrowser.CoreWebView2.DocumentTitle & vbCrLf)
                    'selection.TypeText("URL: " & ChatBrowser.CoreWebView2.Source & " " & "时间: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & vbCrLf)
                    'selection.TypeText("----------------------------------------" & vbCrLf & vbCrLf)

                    'MessageBox.Show("内容已成功提取并插入到文档中", "成功")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"处理提取内容时出错: {ex.Message}", "错误",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 添加视图销毁处理
    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        isViewInitialized = False
        MyBase.OnHandleDestroyed(e)
    End Sub

End Class