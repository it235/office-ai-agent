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

        ' 订阅 SelectionChange 事件 - 使用新的重载方法
        AddHandler Globals.ThisAddIn.Application.SheetSelectionChange, AddressOf GetSelectionContentExcel

    End Sub

    ' 保持原有的Override方法以兼容基类
    Protected Overrides Sub GetSelectionContent(target As Object)
        ' 如果是从Excel的SheetSelectionChange事件调用，target应该是Worksheet
        If TypeOf target Is Microsoft.Office.Interop.Excel.Worksheet Then
            ' 获取当前选中的范围
            Dim selection = Globals.ThisAddIn.Application.Selection
            If TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                GetSelectionContentExcel(target, DirectCast(selection, Microsoft.Office.Interop.Excel.Range))
            End If
        End If
    End Sub

    ' 添加一个新的重载方法来处理Excel的事件
    Private Sub GetSelectionContentExcel(Sh As Microsoft.Office.Interop.Excel.Worksheet, Target As Microsoft.Office.Interop.Excel.Range)
        If Me.Visible AndAlso selectedCellChecked Then
            Dim sheetName As String = Sh.Name
            Dim address As String = Target.Address(False, False)
            Dim key As String = $"{sheetName}"

            ' 添加新的项
            AddSelectedContentItem(key, address)
        End If
    End Sub

    Private Async Sub AddSelectedContentItem(sheetName As String, address As String)
        'Dim ctrlKey As Boolean = False
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control

        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
        InitializeSettings()
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

    Private Function GetSelectedRangeContent() As String
        Try
            ' 获取 sheetContentItems 的内容
            Dim selectedContents As String = String.Join("|", sheetContentItems.Values.Select(Function(item) item.Item1.Text))

            ' 解析 selectedContents 并获取每个工作表中选定的单元格内容
            Dim parsedContents As New StringBuilder()
            If Not String.IsNullOrEmpty(selectedContents) Then
                Dim sheetSelections = selectedContents.Split("|"c)
                For Each sheetSelection In sheetSelections
                    Dim parts = sheetSelection.Split("["c)
                    If parts.Length = 2 Then
                        Dim sheetName = parts(0)
                        Dim ranges = parts(1).TrimEnd("]"c).Split(","c)
                        For Each range In ranges
                            Dim content = GetRangeContent(sheetName, range)
                            If Not String.IsNullOrEmpty(content) Then
                                parsedContents.AppendLine($"{sheetName}的{range}:{content}")
                            End If
                        Next
                    End If
                Next
            End If

            ' 将 parsedContents 加入到 question 中
            If parsedContents.Length > 0 Then
                Return "我能提供我选中的数据作为参考：{" & parsedContents.ToString() & "}"
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Private Function GetRangeContent(sheetName As String, rangeAddress As String) As String
        Try
            Dim sheet = Globals.ThisAddIn.Application.Sheets(sheetName)
            Dim range = sheet.Range(rangeAddress)
            Dim value = range.Value2

            If value Is Nothing Then
                Return String.Empty
            End If

            If TypeOf value Is System.Object(,) Then
                Dim array = DirectCast(value, System.Object(,))
                Dim rows = array.GetLength(0)
                Dim cols = array.GetLength(1)
                Dim result As New StringBuilder()

                For i = 1 To rows
                    For j = 1 To cols
                        If array(i, j) IsNot Nothing Then
                            result.Append(array(i, j).ToString() & vbTab)
                        End If
                    Next
                    result.AppendLine()
                Next

                Return result.ToString().TrimEnd()
            Else
                Return value.ToString()
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function
    Protected Overrides Sub SendChatMessage(message As String)
        ' 这里可以实现word的特殊逻辑

        If selectedCellChecked Then
            ' 获取 sheetContentItems 的内容
            Dim selectedContents As String = String.Join(", ", sheetContentItems.Values.Select(Function(item) item.Item1.Text))
            ' 将 selectedContents 加入到 message 中
            If Not String.IsNullOrEmpty(selectedContents) Then
                message &= " 我能提供我选中的数据作为参考：" & selectedContents
            End If
        End If
        Send(message)
    End Sub


End Class

