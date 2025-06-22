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

            ' 检查选中范围的单元格数量
            Dim cellCount As Integer = Target.Cells.Count

            ' 如果选择了多个单元格，总是添加为引用，不管是否有内容
            If cellCount > 1 Then
                AddSelectedContentItem(key, address)
            Else
                ' 只有单个单元格时，才检查是否有内容
                Dim hasContent As Boolean = False
                For Each cell As Microsoft.Office.Interop.Excel.Range In Target
                    If cell.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(cell.Value.ToString()) Then
                        hasContent = True
                        Exit For
                    End If
                Next

                If hasContent Then
                    ' 选中单元格有内容，添加新的项
                    AddSelectedContentItem(key, address)
                Else
                    ' 选中没有内容，清除相同 sheetName 的引用
                    ClearSelectedContentBySheetName(key)
                End If
            End If
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


    ' 执行 JavaScript 代码，支持操作Excel对象
    Protected Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(jsCode, preview) Then
                    Return False
                End If
            End If

            ' 检查代码类型 - 普通JS还是Excel操作JS
            Dim isExcelJS As Boolean = jsCode.Contains("Excel.") OrElse
                                  jsCode.Contains("ActiveXObject") OrElse
                                  jsCode.Contains("Application") OrElse
                                  jsCode.Contains("Workbook")

            If isExcelJS Then
                ' 创建脚本控制引擎来执行操作Excel的JavaScript
                Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
                scriptEngine.Language = "JScript"

                ' 设置对Excel应用程序的引用
                scriptEngine.AddObject("excelApp", Globals.ThisAddIn.Application, True)

                ' 构建执行代码
                Dim scriptCode As String =
                "function executeExcelJS() {" & vbCrLf &
                "  try {" & vbCrLf &
                "    // Excel已作为excelApp对象提供" & vbCrLf &
                "    " & jsCode & vbCrLf &
                "    return 'JS代码执行成功';" & vbCrLf &
                "  } catch(e) {" & vbCrLf &
                "    return 'JS执行错误: ' + e.message;" & vbCrLf &
                "  }" & vbCrLf &
                "}" & vbCrLf &
                "executeExcelJS();"

                ' 执行JavaScript代码
                Dim result As String = scriptEngine.Eval(scriptCode)
                GlobalStatusStrip.ShowInfo(result)
                Return True
            Else
                ' 对于普通JavaScript，使用WebView2执行
                Dim scriptResult As Task(Of String) = ChatBrowser.ExecuteScriptAsync(jsCode)
                scriptResult.Wait() ' 等待执行完成

                ' 显示结果
                If Not String.IsNullOrEmpty(scriptResult.Result) Then
                    Dim resultStr As String = scriptResult.Result.Trim(""""c) ' 移除JSON字符串引号
                    GlobalStatusStrip.ShowInfo("JS执行结果: " + resultStr)
                End If

                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("执行JavaScript代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
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


    ' 执行SQL查询
    Protected Function ExecuteSqlQuery(sqlCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(sqlCode, preview) Then
                    Return False
                End If
            End If

            ' 获取应用程序信息
            Dim appInfo As ApplicationInfo = GetApplication()

            Dim activeWorkbook As Object = Globals.ThisAddIn.Application.ActiveWorkbook

            ' 创建查询表
            Dim activeSheet As Object = Globals.ThisAddIn.Application.ActiveSheet
            Dim queryTable As Object = Nothing

                ' 获取可用的单元格区域
                Dim targetCell As Object = activeSheet.Range("A1")

                ' 创建SQL连接字符串 (示例使用当前工作簿作为数据源)
                Dim connString As String = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                      activeWorkbook.FullName & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

                ' 创建查询定义
                queryTable = activeSheet.QueryTables.Add(connString, targetCell, sqlCode)

                ' 设置查询属性
                queryTable.RefreshStyle = 1 ' xlOverwriteCells
                queryTable.BackgroundQuery = False

                ' 执行查询
                queryTable.Refresh(False)

                GlobalStatusStrip.ShowWarning("SQL查询已执行")
            Return True
        Catch ex As Exception
            MessageBox.Show("执行SQL查询时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 执行PowerQuery/M语言
    Protected Function ExecutePowerQuery(mCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(mCode, preview) Then
                    Return False
                End If
            End If

            ' 获取应用程序信息
            Dim appInfo As ApplicationInfo = GetApplication()

            ' PowerQuery执行需要较复杂的实现，这里仅提供基本框架
            Dim excelApp = Globals.ThisAddIn.Application
                Dim wb As Object = excelApp.ActiveWorkbook

                ' 检查Excel版本是否支持PowerQuery
                Dim versionSupported As Boolean = excelApp.Version >= 15 ' Excel 2013及以上版本

                If Not versionSupported Then
                    GlobalStatusStrip.ShowWarning("PowerQuery需要Excel 2013或更高版本")
                    Return False
                End If

                ' PowerQuery执行逻辑需要根据具体需求实现
                GlobalStatusStrip.ShowWarning("PowerQuery代码执行功能正在开发中")
            Return True
        Catch ex As Exception
            MessageBox.Show("执行PowerQuery代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 执行Python代码
    Protected Function ExecutePython(pythonCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(pythonCode, preview) Then
                    Return False
                End If
            End If

            ' 获取应用程序信息
            Dim appInfo As ApplicationInfo = GetApplication()

            Dim excelApp = Globals.ThisAddIn.Application

                ' 检查Excel版本是否支持Python (Excel 365)
                Dim versionSupported As Boolean = False

                Try
                    ' 尝试访问Python对象，如果不支持会抛出异常
                    Dim pythonObj As Object = excelApp.PythonExecute("print('test')")
                    versionSupported = True
                Catch
                    versionSupported = False
                End Try

                If Not versionSupported Then
                    ' 如果内置Python不可用，可以尝试通过外部Python解释器执行
                    GlobalStatusStrip.ShowWarning("此Excel版本不支持内置Python，尝试使用外部Python...")

                    ' 创建临时Python文件
                    Dim tempFile As String = Path.Combine(Path.GetTempPath(), "excel_python_" & Guid.NewGuid().ToString() & ".py")
                    File.WriteAllText(tempFile, pythonCode)

                    ' 使用Process类执行Python脚本
                    Dim startInfo As New ProcessStartInfo With {
                    .FileName = "python", ' 假设Python已安装并在PATH中
                    .Arguments = tempFile,
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }

                    Using process As Process = Process.Start(startInfo)
                        Dim output As String = process.StandardOutput.ReadToEnd()
                        Dim error1 As String = process.StandardError.ReadToEnd()
                        process.WaitForExit()

                        If Not String.IsNullOrEmpty(error1) Then
                        GlobalStatusStrip.ShowWarning("Python执行错误: " & error1)
                    Else
                            GlobalStatusStrip.ShowWarning("Python执行结果: " & output)
                        End If
                    End Using

                    ' 删除临时文件
                    Try
                        File.Delete(tempFile)
                    Catch
                        ' 忽略清理错误
                    End Try
                Else
                    ' 使用Excel内置Python执行代码
                    Dim result As Object = excelApp.PythonExecute(pythonCode)
                    GlobalStatusStrip.ShowWarning("Python代码已执行")
                End If

            Return True
        Catch ex As Exception
            MessageBox.Show("执行Python代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        Debug.Print(message)
        Send(message)
    End Sub

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' 获取当前活动工作表和选择区域
            Dim activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim selection = Globals.ThisAddIn.Application.Selection

            ' 如果有选择区域且为 Range 类型
            If selection IsNot Nothing AndAlso TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                Dim selectedRange As Microsoft.Office.Interop.Excel.Range = DirectCast(selection, Microsoft.Office.Interop.Excel.Range)

                ' 创建内容构建器，按照 ParseFile 的结构
                Dim contentBuilder As New StringBuilder()
                contentBuilder.AppendLine(vbCrLf & "--- 用户选中的WorkbookSheet参考内容如下 ---")

                ' 添加活动工作簿信息
                contentBuilder.AppendLine($"工作簿: {Path.GetFileName(activeWorkbook.FullName)}")

                ' 获取选择的工作表信息
                Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = selectedRange.Worksheet
                Dim sheetName As String = worksheet.Name

                ' 添加工作表信息
                contentBuilder.AppendLine($"工作表: {sheetName}")

                ' 获取选择区域的范围地址
                Dim address As String = selectedRange.Address(False, False)
                contentBuilder.AppendLine($"  使用范围: {address}")

                ' 读取选择区域中的单元格内容
                Dim usedRange As Microsoft.Office.Interop.Excel.Range = selectedRange

                ' 获取区域的行列信息
                Dim firstRow As Integer = usedRange.Row
                Dim firstCol As Integer = usedRange.Column
                Dim lastRow As Integer = firstRow + usedRange.Rows.Count - 1
                Dim lastCol As Integer = firstCol + usedRange.Columns.Count - 1

                ' 限制读取的单元格数量（防止数据过大）
                Dim maxRows As Integer = Math.Min(lastRow, firstRow + 30)
                Dim maxCols As Integer = Math.Min(lastCol, firstCol + 10)

                ' 逐个单元格读取内容
                For rowIndex As Integer = firstRow To maxRows
                    For colIndex As Integer = firstCol To maxCols
                        Try
                            Dim cell As Microsoft.Office.Interop.Excel.Range = worksheet.Cells(rowIndex, colIndex)
                            Dim cellValue As Object = cell.Value

                            If cellValue IsNot Nothing Then
                                Dim cellAddress As String = $"{GetExcelColumnName(colIndex)}{rowIndex}"
                                contentBuilder.AppendLine($"  {cellAddress}: {cellValue}")
                            End If
                        Catch cellEx As Exception
                            Debug.WriteLine($"读取单元格时出错: {cellEx.Message}")
                            ' 继续处理下一个单元格
                        End Try
                    Next
                Next

                ' 如果有更多行或列未显示，添加提示
                If lastRow > maxRows Then
                    contentBuilder.AppendLine($"  ... 共有 {lastRow - firstRow + 1} 行，仅显示前 {maxRows - firstRow + 1} 行")
                End If
                If lastCol > maxCols Then
                    contentBuilder.AppendLine($"  ... 共有 {lastCol - firstCol + 1} 列，仅显示前 {maxCols - firstCol + 1} 列")
                End If

                contentBuilder.AppendLine("--- WorkbookSheet参考内容到这结束 ---" & vbCrLf)

                ' 将选中内容添加到消息中
                message &= contentBuilder.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取选中单元格内容时出错: {ex.Message}")
            ' 出错时不添加选中内容，继续发送原始消息
        End Try
        Return message
    End Function

    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 创建一个新的 Excel 应用程序实例（为避免影响当前工作簿）
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Try
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly:=True)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"文件: {Path.GetFileName(filePath)} 包含以下内容:")

                ' 处理每个工作表
                For Each worksheet As Microsoft.Office.Interop.Excel.Worksheet In workbook.Worksheets
                    Dim sheetName As String = worksheet.Name
                    contentBuilder.AppendLine($"工作表: {sheetName}")

                    ' 获取使用范围
                    Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
                    If usedRange IsNot Nothing Then
                        Dim lastRow As Integer = usedRange.Row + usedRange.Rows.Count - 1
                        Dim lastCol As Integer = usedRange.Column + usedRange.Columns.Count - 1

                        ' 限制读取的单元格数量（防止文件过大）
                        Dim maxRows As Integer = Math.Min(lastRow, 30)
                        Dim maxCols As Integer = Math.Min(lastCol, 10)

                        contentBuilder.AppendLine($"  使用范围: {GetExcelColumnName(usedRange.Column)}{usedRange.Row}:{GetExcelColumnName(lastCol)}{lastRow}")

                        ' 读取单元格内容
                        For rowIndex As Integer = usedRange.Row To maxRows
                            For colIndex As Integer = usedRange.Column To maxCols
                                Try
                                    Dim cell As Microsoft.Office.Interop.Excel.Range = worksheet.Cells(rowIndex, colIndex)
                                    Dim cellValue As Object = cell.Value

                                    If cellValue IsNot Nothing Then
                                        Dim cellAddress As String = $"{GetExcelColumnName(colIndex)}{rowIndex}"
                                        contentBuilder.AppendLine($"  {cellAddress}: {cellValue}")
                                    End If
                                Catch cellEx As Exception
                                    Debug.WriteLine($"读取单元格时出错: {cellEx.Message}")
                                    ' 继续处理下一个单元格
                                End Try
                            Next
                        Next

                        ' 如果有更多行或列未显示，添加提示
                        If lastRow > maxRows Then
                            contentBuilder.AppendLine($"  ... 共有 {lastRow - usedRange.Row + 1} 行，仅显示前 {maxRows - usedRange.Row + 1} 行")
                        End If
                        If lastCol > maxCols Then
                            contentBuilder.AppendLine($"  ... 共有 {lastCol - usedRange.Column + 1} 列，仅显示前 {maxCols - usedRange.Column + 1} 列")
                        End If
                    End If

                    contentBuilder.AppendLine()
                Next

                Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Excel",
                .ParsedContent = contentBuilder.ToString(),
                .RawData = Nothing ' 可以选择存储更多数据供后续处理
            }

            Finally
                ' 清理资源
                If workbook IsNot Nothing Then
                    workbook.Close(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                End If

                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析 Excel 文件时出错: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Excel",
            .ParsedContent = $"[解析 Excel 文件时出错: {ex.Message}]"
        }
        End Try
    End Function

    ' 辅助方法：将列索引转换为 Excel 列名（如 1->A, 27->AA）
    Private Function GetExcelColumnName(columnIndex As Integer) As String
        Dim dividend As Integer = columnIndex
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Chr(65 + modulo) & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    ' 实现获取当前 Excel 工作目录的方法
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
End Class

