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


    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
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
        Send(message, "", True, "")
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（Excel特定实现）
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            ' 使用意图优化的systemPrompt
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"Excel使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            Send(message, "", True, "")
        End If
    End Sub

    ''' <summary>
    ''' 获取选中内容并格式化（使用ExcelContextService优化）
    ''' </summary>
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' 获取当前活动工作表和选择区域
            Dim activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim selection = Globals.ThisAddIn.Application.Selection

            ' 如果有选择区域且为 Range 类型
            If selection IsNot Nothing AndAlso TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                Dim selectedRange As Microsoft.Office.Interop.Excel.Range = DirectCast(selection, Microsoft.Office.Interop.Excel.Range)

                ' 提取数据到数组
                Dim data As Object(,) = ExtractRangeData(selectedRange)

                If data IsNot Nothing AndAlso data.Length > 0 Then
                    ' 使用ExcelContextService进行优化的数据格式化
                    Dim contextService As New ShareRibbon.ExcelContextService()
                    Dim workbookName As String = Path.GetFileName(activeWorkbook.FullName)
                    Dim worksheetName As String = selectedRange.Worksheet.Name
                    Dim rangeAddress As String = selectedRange.Address(False, False)

                    ' 调用优化的格式化方法
                    Dim formattedContent As String = contextService.FormatSelectionAsContext(
                        data,
                        workbookName,
                        worksheetName,
                        rangeAddress)

                    If Not String.IsNullOrEmpty(formattedContent) Then
                        message &= formattedContent
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取选中单元格内容时出错: {ex.Message}")
            ' 出错时不添加选中内容，继续发送原始消息
        End Try
        Return message
    End Function

    ''' <summary>
    ''' 从Excel Range提取数据到二维数组（高性能批量读取）
    ''' </summary>
    Private Function ExtractRangeData(selectedRange As Microsoft.Office.Interop.Excel.Range) As Object(,)
        Try
            Const MAX_ROWS As Integer = 100
            Const MAX_COLS As Integer = 26

            ' 限制读取范围
            Dim rowCount = Math.Min(selectedRange.Rows.Count, MAX_ROWS)
            Dim colCount = Math.Min(selectedRange.Columns.Count, MAX_COLS)

            ' 如果范围过大，只取部分
            Dim actualRange As Microsoft.Office.Interop.Excel.Range
            If selectedRange.Rows.Count > MAX_ROWS OrElse selectedRange.Columns.Count > MAX_COLS Then
                actualRange = selectedRange.Resize(rowCount, colCount)
            Else
                actualRange = selectedRange
            End If

            ' 一次性读取所有数据（性能优化关键）
            Dim values As Object = actualRange.Value2

            ' 处理单个单元格的情况
            If Not TypeOf values Is Object(,) Then
                ' 创建1-based数组（与Excel返回的多单元格数组一致）
                Dim result = DirectCast(Array.CreateInstance(GetType(Object), {1, 1}, {1, 1}), Object(,))
                result(1, 1) = values
                Return result
            End If

            Return DirectCast(values, Object(,))

        Catch ex As Exception
            Debug.WriteLine($"ExtractRangeData 出错: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 尝试执行AI返回的直接操作命令
    ''' </summary>
    ''' <param name="aiResponse">AI响应文本</param>
    ''' <returns>是否成功执行了命令</returns>
    Public Function TryExecuteDirectCommands(aiResponse As String) As Boolean
        Try
            ' 提取JSON命令
            Dim commands = ExcelDirectOperationService.ExtractCommandsFromResponse(aiResponse)

            If commands.Count = 0 Then
                Return False
            End If

            ' 创建操作服务
            Dim operationService As New ExcelDirectOperationService(Globals.ThisAddIn.Application)

            ' 执行所有命令
            Dim allSuccess As Boolean = True
            For Each cmd As Newtonsoft.Json.Linq.JObject In commands
                Dim success = operationService.ExecuteCommand(cmd)
                If Not success Then
                    allSuccess = False
                End If
            Next

            Return allSuccess

        Catch ex As Exception
            Debug.WriteLine($"TryExecuteDirectCommands 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带占位符替换和校验
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommand(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 获取Excel上下文用于占位符替换
            Dim context = ExcelJsonCommandSchema.GetExcelContext(Globals.ThisAddIn.Application)
            
            ' 先替换占位符再解析JSON
            Dim processedJson = jsonCode
            For Each kvp In context
                processedJson = processedJson.Replace("{" & kvp.Key & "}", kvp.Value)
            Next
            
            ' 使用严格的结构验证
            Dim errorMessage As String = ""
            Dim normalizedJson As Newtonsoft.Json.Linq.JToken = Nothing
            
            If Not ExcelJsonCommandSchema.ValidateJsonStructure(processedJson, errorMessage, normalizedJson) Then
                ' 格式验证失败，显示详细错误
                Debug.WriteLine($"JSON格式验证失败: {errorMessage}")
                Debug.WriteLine($"原始JSON: {processedJson.Substring(0, Math.Min(200, processedJson.Length))}...")
                
                ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式不符合规范: {errorMessage}")
                
                ' 通知前端显示格式修正提示
                Dim correctionPrompt = ExcelJsonCommandSchema.GetFormatCorrectionPrompt(
                    processedJson.Substring(0, Math.Min(500, processedJson.Length)),
                    errorMessage)
                Debug.WriteLine($"格式修正提示已生成，长度: {correctionPrompt.Length}")
                
                Return False
            End If
            
            ' 验证通过，根据类型执行
            If normalizedJson.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, Newtonsoft.Json.Linq.JObject)
                
                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecuteCommandsArray(jsonObj("commands"), processedJson, preview, context)
                End If
                
                ' 单命令格式
                Return ExecuteSingleCommand(jsonObj, processedJson, preview)
            End If
            
            ShareRibbon.GlobalStatusStrip.ShowWarning("无效的JSON格式")
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式无效: {ex.Message}")
            Return False
        Catch ex As Exception
            ShareRibbon.GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 将 actions 数组转换为标准的 commands 数组格式
    ''' <summary>
    ''' 执行命令数组
    ''' </summary>
    Private Function ExecuteCommandsArray(commandsArray As Newtonsoft.Json.Linq.JToken, originalJson As String, preview As Boolean, context As Dictionary(Of String, String)) As Boolean
        Try
            Dim commands = CType(commandsArray, Newtonsoft.Json.Linq.JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("命令数组为空")
                Return False
            End If

            ' 创建操作服务
            Dim operationService As New ExcelDirectOperationService(Globals.ThisAddIn.Application)

            ' 预览所有命令 - 使用 JsonPreviewDialog
            If preview Then
                Try
                    ' 生成批量命令的预览结果
                    Dim previewResult = GenerateBatchPreviewResult(commands, originalJson)
                    
                    ' 显示预览对话框
                    Using dialog As New JsonPreviewDialog()
                        If dialog.ShowPreview(previewResult) <> DialogResult.OK Then
                            ShareRibbon.GlobalStatusStrip.ShowInfo("用户取消执行")
                            ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                            Return True
                        End If
                    End Using
                Catch previewEx As Exception
                    Debug.WriteLine($"批量预览生成失败，使用简单预览: {previewEx.Message}")
                    ' 回退到简单预览
                    Dim previewMsg As New StringBuilder()
                    previewMsg.AppendLine($"即将执行 {commands.Count} 个命令:")
                    previewMsg.AppendLine()

                    Dim cmdIndex = 1
                    For Each cmd In commands
                        If cmd.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim cmdObj = CType(cmd, Newtonsoft.Json.Linq.JObject)
                            Dim cmdName = cmdObj("command")?.ToString()
                            Dim range = If(cmdObj("range")?.ToString(), cmdObj("params")?("range")?.ToString())
                            Dim formula = If(cmdObj("formula")?.ToString(), cmdObj("params")?("formula")?.ToString())
                            
                            previewMsg.AppendLine($"{cmdIndex}. {cmdName}")
                            If Not String.IsNullOrEmpty(range) Then previewMsg.AppendLine($"   范围: {range}")
                            If Not String.IsNullOrEmpty(formula) Then previewMsg.AppendLine($"   公式: {formula}")
                            previewMsg.AppendLine()
                            cmdIndex += 1
                        End If
                    Next

                    previewMsg.AppendLine("是否继续执行？")

                    If MessageBox.Show(previewMsg.ToString(), "批量命令预览", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> DialogResult.OK Then
                        ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                        Return True
                    End If
                End Try
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0
            Dim cmdNumber = 1

            For Each cmd In commands
                If cmd.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                    Dim cmdObj = CType(cmd, Newtonsoft.Json.Linq.JObject)
                    
                    ' 标准化命令结构
                    cmdObj = ExcelJsonCommandSchema.NormalizeCommandStructure(cmdObj)
                    
                    ' 校验命令
                    Dim errorMsg As String = ""
                    If Not ExcelJsonCommandSchema.ValidateCommand(cmdObj, errorMsg) Then
                        Debug.WriteLine($"命令 {cmdNumber} 校验失败: {errorMsg}")
                        failCount += 1
                        cmdNumber += 1
                        Continue For
                    End If

                    ' 执行命令
                    If operationService.ExecuteCommand(cmdObj) Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                End If
                cmdNumber += 1
            Next

            If failCount = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"所有 {successCount} 个命令执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"执行完成: {successCount} 成功, {failCount} 失败")
            End If

            Return failCount = 0

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"批量执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 生成批量命令的预览结果
    ''' </summary>
    Private Function GenerateBatchPreviewResult(commands As Newtonsoft.Json.Linq.JArray, originalJson As String) As JsonPreviewResult
        Dim result As New JsonPreviewResult()
        result.OriginalJson = originalJson

        ' 生成执行计划
        result.ExecutionPlan = New List(Of ExecutionStep)()
        Dim stepNumber = 1

        For Each cmd In commands
            If cmd.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim cmdObj = CType(cmd, Newtonsoft.Json.Linq.JObject)
                Dim cmdName = cmdObj("command")?.ToString()
                Dim range = If(cmdObj("range")?.ToString(), 
                            If(cmdObj("params")?("range")?.ToString(),
                            If(cmdObj("params")?("targetRange")?.ToString(), "")))
                Dim formula = If(cmdObj("formula")?.ToString(), cmdObj("params")?("formula")?.ToString())

                Dim stepIcon = GetCommandIcon(cmdName)
                Dim stepDesc = GetCommandDescription(cmdName, formula, range)

                result.ExecutionPlan.Add(New ExecutionStep(stepNumber, stepDesc, stepIcon) With {
                    .WillModify = range
                })
                stepNumber += 1
            End If
        Next

        ' 生成摘要
        Dim summaryBuilder As New StringBuilder()
        summaryBuilder.AppendLine($"即将执行 {commands.Count} 个命令")
        summaryBuilder.AppendLine()
        summaryBuilder.AppendLine("命令列表:")
        
        Dim cmdIndex = 1
        For Each cmd In commands
            If cmd.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim cmdObj = CType(cmd, Newtonsoft.Json.Linq.JObject)
                Dim cmdName = cmdObj("command")?.ToString()
                Dim formula = If(cmdObj("formula")?.ToString(), cmdObj("params")?("formula")?.ToString())
                Dim range = If(cmdObj("range")?.ToString(), cmdObj("params")?("range")?.ToString())
                
                summaryBuilder.AppendLine($"  {cmdIndex}. {cmdName}")
                If Not String.IsNullOrEmpty(formula) Then summaryBuilder.AppendLine($"      公式: {formula}")
                If Not String.IsNullOrEmpty(range) Then summaryBuilder.AppendLine($"      范围: {range}")
                cmdIndex += 1
            End If
        Next

        result.Summary = summaryBuilder.ToString()

        Return result
    End Function

    ''' <summary>
    ''' 获取命令对应的图标类型
    ''' </summary>
    Private Function GetCommandIcon(command As String) As String
        Select Case command?.ToLower()
            Case "applyformula"
                Return "formula"
            Case "writedata"
                Return "data"
            Case "formatrange"
                Return "format"
            Case "createchart"
                Return "chart"
            Case "cleandata"
                Return "clean"
            Case Else
                Return "default"
        End Select
    End Function

    ''' <summary>
    ''' 获取命令描述
    ''' </summary>
    Private Function GetCommandDescription(command As String, formula As String, range As String) As String
        Select Case command?.ToLower()
            Case "applyformula"
                If Not String.IsNullOrEmpty(formula) Then
                    Return $"应用公式 {formula}"
                End If
                Return "应用公式"
            Case "writedata"
                Return "写入数据"
            Case "formatrange"
                Return "设置格式"
            Case "createchart"
                Return "创建图表"
            Case "cleandata"
                Return "清洗数据"
            Case Else
                Return command
        End Select
    End Function

    ''' <summary>
    ''' 执行单个命令
    ''' </summary>
    Private Function ExecuteSingleCommand(commandJson As Newtonsoft.Json.Linq.JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            
            ' 校验JSON命令
            Dim errorMsg As String = ""
            If Not ExcelJsonCommandSchema.ValidateCommand(commandJson, errorMsg) Then
                ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON命令格式错误: {errorMsg}")
                Return False
            End If

            ' 创建操作服务
            Dim operationService As New ExcelDirectOperationService(Globals.ThisAddIn.Application)

            ' JSON预览对话框
            If preview Then
                Try
                    ' 生成预览结果
                    Dim previewResult = GenerateJsonPreviewResult(commandJson, processedJson, operationService)
                    
                    ' 显示预览对话框
                    Using dialog As New JsonPreviewDialog()
                        If dialog.ShowPreview(previewResult) <> DialogResult.OK Then
                            ShareRibbon.GlobalStatusStrip.ShowInfo("用户取消执行")
                            ' 通知前端用户取消了执行（恢复按钮可点击状态）
                            ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                            Return True ' 返回True表示正常取消，而非错误
                        End If
                    End Using
                Catch previewEx As Exception
                    Debug.WriteLine($"预览生成失败，使用简单预览: {previewEx.Message}")
                    ' 回退到简单预览
                    Dim params = commandJson("params")
                    Dim targetRange = If(params?("targetRange")?.ToString(), params?("range")?.ToString())
                    Dim formula = params?("formula")?.ToString()

                    Dim previewMsg = $"即将执行 Excel 命令:{vbCrLf}{vbCrLf}" &
                                    $"命令: {command}{vbCrLf}" &
                                    If(Not String.IsNullOrEmpty(targetRange), $"目标: {targetRange}{vbCrLf}", "") &
                                    If(Not String.IsNullOrEmpty(formula), $"公式: {formula}{vbCrLf}", "") &
                                    $"{vbCrLf}是否继续执行？"

                    If MessageBox.Show(previewMsg, "JSON命令预览", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> DialogResult.OK Then
                        ' 通知前端用户取消了执行
                        ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                        Return True
                    End If
                End Try
            End If

            ' 执行命令
            Dim success = operationService.ExecuteCommand(commandJson)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"命令 '{command}' 执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"命令 '{command}' 执行失败")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 生成JSON命令预览结果
    ''' </summary>
    Private Function GenerateJsonPreviewResult(commandJson As JObject, originalJson As String, operationService As ExcelDirectOperationService) As JsonPreviewResult
        Dim result As New JsonPreviewResult()
        result.OriginalJson = originalJson

        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")

            ' 使用 ExecutionPlanRenderer 生成执行计划
            Dim renderer As New ShareRibbon.ExecutionPlanRenderer()
            result.ExecutionPlan = renderer.ParseJsonToExecutionPlan(originalJson)

            ' 生成摘要
            Dim summaryBuilder As New StringBuilder()
            summaryBuilder.AppendLine($"将执行 {command} 命令")

            Dim targetRange = If(params?("targetRange")?.ToString(), params?("range")?.ToString())
            If Not String.IsNullOrEmpty(targetRange) Then
                summaryBuilder.AppendLine($"目标范围: {targetRange}")
            End If

            Dim formula = params?("formula")?.ToString()
            If Not String.IsNullOrEmpty(formula) Then
                summaryBuilder.AppendLine($"公式: {formula}")
            End If

            result.Summary = summaryBuilder.ToString()

            ' 尝试预测单元格变更（简化实现，仅显示目标范围）
            If Not String.IsNullOrEmpty(targetRange) Then
                result.CellChanges = New List(Of CellChange)()
                
                ' 简化的变更预测：标记目标范围会被修改
                result.CellChanges.Add(New CellChange() With {
                    .Address = targetRange,
                    .ChangeType = "Modified",
                    .OldValue = "(当前值)",
                    .NewValue = "(新值)"
                })
            End If

        Catch ex As Exception
            Debug.WriteLine($"GenerateJsonPreviewResult 出错: {ex.Message}")
            result.Summary = "无法生成详细预览"
        End Try

        Return result
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


    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)

    End Sub

    ''' <summary>
    ''' 获取当前Excel上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "Excel"

        Try
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection IsNot Nothing AndAlso TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                Dim selRange = DirectCast(selection, Microsoft.Office.Interop.Excel.Range)

                ' 获取选中区域地址
                Dim rangeAddr = ""
                Try
                    rangeAddr = selRange.Address(False, False)
                Catch
                End Try
                snapshot("selectionAddress") = rangeAddr

                ' 获取选中内容（限制大小）
                Dim selText = ""
                Try
                    Dim cellValue = selRange.Value2
                    If cellValue IsNot Nothing Then
                        If TypeOf cellValue Is Object(,) Then
                            Dim values = DirectCast(cellValue, Object(,))
                            Dim sb As New StringBuilder()

                            ' 使用正确的数组边界（Excel返回1-based数组）
                            Dim rowStart = values.GetLowerBound(0)
                            Dim rowEnd = values.GetUpperBound(0)
                            Dim colStart = values.GetLowerBound(1)
                            Dim colEnd = values.GetUpperBound(1)

                            Dim maxRowEnd = Math.Min(rowEnd, rowStart + 4)
                            Dim maxColEnd = Math.Min(colEnd, colStart + 4)

                            For r = rowStart To maxRowEnd
                                For c = colStart To maxColEnd
                                    If values(r, c) IsNot Nothing Then
                                        sb.Append(values(r, c).ToString())
                                    End If
                                    If c < maxColEnd Then sb.Append(vbTab)
                                Next
                                sb.AppendLine()
                            Next
                            selText = sb.ToString()
                        Else
                            selText = cellValue.ToString()
                        End If
                    End If
                Catch
                End Try

                If selText.Length > 300 Then
                    selText = selText.Substring(0, 300) & "..."
                End If
                snapshot("selection") = selText
            Else
                snapshot("selection") = ""
            End If

            ' 获取工作表名
            Dim ws = Globals.ThisAddIn.Application.ActiveSheet
            If ws IsNot Nothing Then
                snapshot("sheetName") = CStr(ws.Name)
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function
End Class

