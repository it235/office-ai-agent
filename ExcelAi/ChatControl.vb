Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports ExcelAi.Extensions
Imports ShareRibbon.Extensions
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl


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

    Private Overloads Async Sub AddSelectedContentItem(sheetName As String, address As String)
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

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function

    Protected Overrides Function GetOfficeAppType() As String
        Return "Excel"
    End Function

    ''' <summary>
    ''' 会话绑定的文档：只在 UI 线程读取活动工作簿完整路径，未保存工作簿按通用会话处理。
    ''' </summary>
    Protected Overrides Function GetSessionDocument() As SessionDocument
        Try
            Dim wb = Globals.ThisAddIn.Application.ActiveWorkbook
            If wb Is Nothing Then Return Nothing
            Return SessionDocument.FromFullPath(wb.FullName)
        Catch ex As Exception
            AppLogger.Debug("ExcelAi", "读取活动工作簿标识失败", ex)
            Return Nothing
        End Try
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

                ' 检查数据是否为空（如果选中区域没有实际内容，不发送给LLM）
                If data Is Nothing OrElse data.Length = 0 Then
                    Debug.WriteLine("Excel选中区域数据为空，跳过发送给LLM")
                    Return message
                End If

                ' 检查数据是否全部为空值
                Dim hasContent As Boolean = False
                For Each item In data
                    If item IsNot Nothing AndAlso Not String.IsNullOrEmpty(item.ToString()) Then
                        hasContent = True
                        Exit For
                    End If
                Next

                If Not hasContent Then
                    Debug.WriteLine("Excel选中区域内容全部为空，跳过发送给LLM")
                    Return message
                End If

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
    ''' 执行JSON命令（重写基类方法）- 带占位符替换和校验
    ''' </summary>
    Private Function ExecuteJsonCommandCore(jsonCode As String, preview As Boolean) As Boolean
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
                Debug.WriteLine($"原始JSON: {processedJson}")

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
                            NotifyPreviewCancelled()
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
                        NotifyPreviewCancelled()
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
                            NotifyPreviewCancelled()
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
                        NotifyPreviewCancelled()
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
        ' 在 AI 操作完成时创建撤销点 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                UndoManagerExtension.CreateAIOperationUndoPoint(
                    "Excel",
                    Globals.ThisAddIn.Application,
                    "AI操作",
                    "AI 生成内容")
            End Sub,
            "ExcelAi.ChatControl",
            "创建撤销点")

        ' 尝试检测并应用公式 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                If allPlainMarkdownBuffer IsNot Nothing AndAlso allPlainMarkdownBuffer.Length > 0 Then
                    Dim aiResponse As String = allPlainMarkdownBuffer.ToString()

                    ' 检测并应用公式
                    If FormulaHandlerExtension.DetectFormula(aiResponse) Then
                        System.Diagnostics.Debug.WriteLine("[ChatControl] 检测到公式，尝试应用")

                        ' 在应用公式前创建更具体的撤销点
                        UndoManagerExtension.CreateAIOperationUndoPoint(
                            "Excel",
                            Globals.ThisAddIn.Application,
                            "AI公式应用",
                            "应用AI生成的Excel公式")

                        FormulaHandlerExtension.TryApplyFormula(aiResponse, Globals.ThisAddIn.Application)

                        ' 显示撤销提示
                        Dim hint = UndoManagerExtension.GetUndoHint("Excel")
                        System.Diagnostics.Debug.WriteLine("[ChatControl] 撤销提示: " & hint)
                    End If
                End If
            End Sub,
            "ExcelAi.ChatControl",
            "公式应用",
            "应用公式时出错，请检查公式格式。")
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

    ''' <summary>
    ''' 捕获 Agent 启动用 Excel 上下文快照
    ''' </summary>
    Protected Overrides Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Try
            Return New Context.ExcelContextProvider(Globals.ThisAddIn.Application).GetContext()
        Catch ex As Exception
            Debug.WriteLine($"CaptureOfficeContext 出错: {ex.Message}")
            Return New Agent.Context.OfficeContext With {.AppType = appType}
        End Try
    End Function

    ''' <summary>
    ''' Agent 原生工具必须返回可观察的 ToolResult。Excel 先以宿主快照覆盖全部 JSON 命令，
    ''' 高频写工具会额外记录目标 Range、公式错误数、UsedRange 与图表数量变化。
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Dim commandEnvelope As JObject = Nothing
        Dim toolId As String = "ExcelCommands"
        Dim beforeSnapshot As JObject = Nothing

        Try
            commandEnvelope = ParseExcelCommandEnvelope(jsonCode)
            toolId = GetEnvelopeToolId(commandEnvelope, "ExcelCommands")
            beforeSnapshot = CaptureExcelCommandSnapshot(commandEnvelope)
            Dim success = ExecuteJsonCommandCore(jsonCode, preview)
            ' 预览被用户拒绝时核心返回 True（UI 语义：取消不是错误），但 ToolResult 不能据此报成功。
            If ConsumePreviewCancelled() Then
                Return Agent.ToolResult.PreviewCancelled(toolId)
            End If
            Dim afterSnapshot = CaptureExcelCommandSnapshot(commandEnvelope)
            Dim observation = BuildExcelCommandObservation(toolId, commandEnvelope, success, beforeSnapshot, afterSnapshot)
            Dim summary = If(observation?("summary")?.ToString(), $"{toolId} 执行完成")
            Dim observationError = If(afterSnapshot?("captureError")?.ToString(),
                                      beforeSnapshot?("captureError")?.ToString())

            ' A write may already have succeeded even when observation fails. Never
            ' repeat it automatically, because a second execution could duplicate data.
            If success AndAlso Not String.IsNullOrWhiteSpace(observationError) Then
                observation("rollback") = BuildExcelRollbackNote(True)
                Return Agent.ToolResult.Failed(
                    toolId,
                    $"Excel 命令可能已执行，但结果观察失败: {observationError}",
                    errorCode:=ExceptionClassifier.CodeObservationFailed,
                    userMessage:="Excel 操作可能已经完成，但插件无法安全验证结果；已停止自动重试",
                    recoverable:=False,
                    observation:=observation)
            End If

            If success Then
                Return Agent.ToolResult.Succeed(toolId,
                                                summary,
                                                data:=New With {.targetRefs = observation("targetRefs")},
                                                observation:=observation)
            End If

            ' 失败但工作表确实被改过：Excel 无法安全自动撤销，只能把半应用状态如实报给 Agent，
            ' 并用 recoverable=False 掐断自动重放——重放会在已改过的表上再改一次。
            Dim partialApply = observation.Value(Of Boolean)("changed")
            observation("rollback") = BuildExcelRollbackNote(partialApply)
            Dim failureMessage = If(partialApply, $"{summary}；工作表已被部分修改且未回滚", summary)
            Return Agent.ToolResult.Failed(toolId,
                                           failureMessage,
                                           errorCode:=If(partialApply,
                                                         ExceptionClassifier.CodePartialApply,
                                                         ExceptionClassifier.CodeUnknown),
                                           userMessage:=If(partialApply,
                                                           "Excel 操作部分完成，插件无法自动撤销，请检查工作表后再继续",
                                                           failureMessage),
                                           recoverable:=Not partialApply,
                                           observation:=observation)
        Catch ex As Exception
            Dim crashResult = Agent.ToolResult.FromException(toolId, ex)
            Dim crashObservation = BuildExcelCrashObservation(commandEnvelope, beforeSnapshot)
            If crashObservation IsNot Nothing Then
                crashResult.Observation = crashObservation
                crashResult.Recoverable = Not crashObservation.Value(Of Boolean)("changed")
            End If
            Return crashResult
        End Try
    End Function

    ''' <summary>
    ''' Excel 没有“一条命令 = 一个撤销记录”的事务边界：Application.Undo 会连带撤掉用户此前的操作，
    ''' 所以宿主只做诚实上报——attempted=False 表示根本没尝试回滚，不让 Agent 误以为文档已还原。
    ''' </summary>
    Private Shared Function BuildExcelRollbackNote(partialApply As Boolean) As JObject
        Return New JObject From {
            {"succeeded", False},
            {"attempted", False},
            {"verified", False},
            {"partialApply", partialApply},
            {"error", If(partialApply, "Excel 不支持安全地自动撤销本轮改动，请人工检查或手动撤销", "本轮未检测到工作表变化，无需回滚")}
        }
    End Function

    Private Function BuildExcelCrashObservation(envelope As JObject, beforeSnapshot As JObject) As JObject
        Try
            If envelope Is Nothing OrElse beforeSnapshot Is Nothing Then Return Nothing
            Dim afterSnapshot = CaptureExcelCommandSnapshot(envelope)
            Dim changed = Not JToken.DeepEquals(beforeSnapshot, afterSnapshot)
            Dim observation As New JObject From {
                {"kind", "write"},
                {"summary", "Excel 命令执行中抛出异常"},
                {"changed", changed},
                {"targetRefs", New JArray("Excel:ActiveSheet")},
                {"before", beforeSnapshot},
                {"after", afterSnapshot},
                {"warnings", New JArray()}
            }
            If changed Then CType(observation("warnings"), JArray).Add("异常发生在工作表已发生变化之后，本轮改动未回滚")
            observation("rollback") = BuildExcelRollbackNote(changed)
            Return observation
        Catch ex As Exception
            Debug.WriteLine($"[Excel] 异常路径观察采集失败: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Function ParseExcelCommandEnvelope(jsonCode As String) As JObject
        If String.IsNullOrWhiteSpace(jsonCode) Then Return Nothing

        Dim processedJson = jsonCode
        Try
            Dim context = ExcelJsonCommandSchema.GetExcelContext(Globals.ThisAddIn.Application)
            For Each kvp In context
                processedJson = processedJson.Replace("{" & kvp.Key & "}", kvp.Value)
            Next
        Catch
        End Try

        Dim token = JToken.Parse(processedJson)
        Return TryCast(token, JObject)
    End Function

    Private Shared Function GetEnvelopeToolId(envelope As JObject, fallback As String) As String
        If envelope Is Nothing Then Return fallback
        If envelope("commands") IsNot Nothing Then Return fallback
        Dim command = envelope("command")?.ToString()
        Return If(String.IsNullOrWhiteSpace(command), fallback, command.Trim())
    End Function

    Private Function CaptureExcelCommandSnapshot(envelope As JObject) As JObject
        Dim snapshot As New JObject()
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        Dim usedRange As Object = Nothing
        Dim chartObjects As Object = Nothing
        Dim selectedRange As Object = Nothing
        Dim target As Object = Nothing
        Try
            ' Keep observer COM values late-bound. WPS/部分 Excel 兼容层可以通过
            ' IDispatch 使用 Range，但不支持嵌入 PIA Range 的 QueryInterface。
            Dim app As Object = Globals.ThisAddIn.Application
            workbook = app.ActiveWorkbook
            worksheet = app.ActiveSheet
            If workbook Is Nothing OrElse worksheet Is Nothing Then Return snapshot

            snapshot("workbook") = If(workbook.Name, "")
            snapshot("worksheet") = If(worksheet.Name, "")
            usedRange = worksheet.UsedRange
            If usedRange IsNot Nothing Then
                snapshot("usedRows") = CInt(usedRange.Rows.Count)
                snapshot("usedColumns") = CInt(usedRange.Columns.Count)
            End If
            chartObjects = worksheet.ChartObjects()
            snapshot("chartCount") = CInt(chartObjects.Count)

            Dim targetAddress = ResolveExcelTargetAddress(envelope)
            If String.IsNullOrWhiteSpace(targetAddress) Then
                selectedRange = app.Selection
                If selectedRange IsNot Nothing Then
                    Try
                        targetAddress = CStr(selectedRange.Address(False, False))
                    Catch
                    End Try
                End If
            End If

            snapshot("targetAddress") = If(targetAddress, "")
            If Not String.IsNullOrWhiteSpace(targetAddress) Then
                target = worksheet.Range(targetAddress)
                snapshot("targetValueHash") = ComputeObservationHash(SerializeExcelRangeValue(target.Value2))
                snapshot("targetFormulaHash") = ComputeObservationHash(SerializeExcelRangeValue(target.Formula))
                snapshot("targetPreview") = BuildExcelRangePreview(target, 3, 4)
                snapshot("formulaErrorCount") = CountExcelFormulaErrors(target, 5000)
            End If
        Catch ex As Exception
            snapshot("captureError") = AppLogger.Redact(ex.Message)
            AppLogger.Warn("ExcelObserver", $"Capture snapshot failed: {AppLogger.Redact(ex.Message)}")
        Finally
            ComObjectHelper.ReleaseComObject(target)
            ComObjectHelper.ReleaseComObject(selectedRange)
            ComObjectHelper.ReleaseComObject(chartObjects)
            ComObjectHelper.ReleaseComObject(usedRange)
            ComObjectHelper.ReleaseComObject(worksheet)
            ComObjectHelper.ReleaseComObject(workbook)
        End Try
        Return snapshot
    End Function

    Private Shared Function ResolveExcelTargetAddress(envelope As JObject) As String
        If envelope Is Nothing Then Return ""
        Dim command = envelope
        If envelope("commands") IsNot Nothing AndAlso envelope("commands").Type = JTokenType.Array Then
            command = TryCast(envelope("commands").FirstOrDefault(), JObject)
        End If
        If command Is Nothing Then Return ""
        Dim params = command("params")
        Return If(command("range")?.ToString(),
                  If(params?("range")?.ToString(), params?("targetRange")?.ToString()))
    End Function

    Private Shared Function BuildExcelCommandObservation(toolId As String,
                                                         envelope As JObject,
                                                         success As Boolean,
                                                         beforeSnapshot As JObject,
                                                         afterSnapshot As JObject) As JObject
        Dim changed = Not JToken.DeepEquals(beforeSnapshot, afterSnapshot)
        Dim targetAddress = If(afterSnapshot?("targetAddress")?.ToString(), beforeSnapshot?("targetAddress")?.ToString())
        Dim targetRefs As New JArray()
        If Not String.IsNullOrWhiteSpace(targetAddress) Then
            targetRefs.Add($"Excel:{afterSnapshot?("worksheet")?.ToString()}!{targetAddress}")
        Else
            targetRefs.Add("Excel:ActiveSheet")
        End If

        Dim warnings As New JArray()
        If success AndAlso Not changed Then warnings.Add("命令已处理，但宿主快照未检测到变化；可能为用户取消、格式等价或 noop")
        If afterSnapshot?("captureError") IsNot Nothing Then warnings.Add(afterSnapshot("captureError"))

        Dim diff As New JObject From {
            {"usedRowsDelta", GetSnapshotInteger(afterSnapshot, "usedRows") - GetSnapshotInteger(beforeSnapshot, "usedRows")},
            {"usedColumnsDelta", GetSnapshotInteger(afterSnapshot, "usedColumns") - GetSnapshotInteger(beforeSnapshot, "usedColumns")},
            {"chartCountDelta", GetSnapshotInteger(afterSnapshot, "chartCount") - GetSnapshotInteger(beforeSnapshot, "chartCount")},
            {"formulaErrorDelta", GetSnapshotInteger(afterSnapshot, "formulaErrorCount") - GetSnapshotInteger(beforeSnapshot, "formulaErrorCount")}
        }

        Return New JObject From {
            {"kind", "write"},
            {"summary", If(success, $"Excel 工具 {toolId} 已执行", $"Excel 工具 {toolId} 执行失败")},
            {"targetRefs", targetRefs},
            {"changed", changed},
            {"before", beforeSnapshot},
            {"after", afterSnapshot},
            {"diff", diff},
            {"warnings", warnings}
        }
    End Function

    Private Shared Function GetSnapshotInteger(snapshot As JObject, name As String) As Integer
        If snapshot Is Nothing Then Return 0
        Return If(snapshot(name)?.Value(Of Integer)(), 0)
    End Function

    Private Shared Function SerializeExcelRangeValue(value As Object) As String
        If value Is Nothing Then Return ""
        Try
            Return JsonConvert.SerializeObject(value, Formatting.None)
        Catch
            Return value.ToString()
        End Try
    End Function

    Private Shared Function ComputeObservationHash(value As String) As String
        Dim bytes = Encoding.UTF8.GetBytes(If(value, ""))
        Using sha = System.Security.Cryptography.SHA256.Create()
            Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Shared Function BuildExcelRangePreview(target As Object,
                                                   maxRows As Integer,
                                                   maxColumns As Integer) As String
        If target Is Nothing Then Return ""
        Dim rowCollection As Object = Nothing
        Dim columnCollection As Object = Nothing
        Dim cells As Object = Nothing
        Try
            rowCollection = target.Rows
            columnCollection = target.Columns
            cells = target.Cells
            Dim rows = Math.Min(CInt(rowCollection.Count), maxRows)
            Dim columns = Math.Min(CInt(columnCollection.Count), maxColumns)
            Dim lines As New List(Of String)()
            For rowIndex = 1 To rows
                Dim values As New List(Of String)()
                For columnIndex = 1 To columns
                    Dim cell As Object = Nothing
                    Try
                        cell = cells.Item(rowIndex, columnIndex)
                        Dim cellValue = cell.Value2
                        values.Add(If(cellValue Is Nothing, "", cellValue.ToString()))
                    Finally
                        ComObjectHelper.ReleaseComObject(cell)
                    End Try
                Next
                lines.Add(String.Join(" | ", values))
            Next
            Dim text = String.Join(vbLf, lines)
            Return If(text.Length <= 400, text, text.Substring(0, 400))
        Finally
            ComObjectHelper.ReleaseComObject(cells)
            ComObjectHelper.ReleaseComObject(columnCollection)
            ComObjectHelper.ReleaseComObject(rowCollection)
        End Try
    End Function

    Private Shared Function CountExcelFormulaErrors(target As Object, maxCells As Integer) As Integer
        If target Is Nothing Then Return 0
        Dim count As Integer = 0
        Dim cells As Object = Nothing
        Try
            cells = target.Cells
            Dim inspected = Math.Min(CInt(cells.Count), maxCells)
            For index = 1 To inspected
                Dim cell As Object = Nothing
                Try
                    cell = cells.Item(index)
                    If Information.IsError(cell.Value) Then count += 1
                Catch
                Finally
                    ComObjectHelper.ReleaseComObject(cell)
                End Try
            Next
            Return count
        Finally
            ComObjectHelper.ReleaseComObject(cells)
        End Try
    End Function

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overrides Function GetOfficeApplicationName() As String
        Return "Excel"
    End Function
End Class

