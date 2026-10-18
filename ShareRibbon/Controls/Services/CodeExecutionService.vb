' ShareRibbon\Controls\Services\CodeExecutionService.vb
' 代码执行服务：VBA、JavaScript、Excel公式执行

Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop

''' <summary>
''' 代码执行服务，负责执行 VBA、JavaScript 和 Excel 公式
''' </summary>
Public Class CodeExecutionService
        Private ReadOnly _getVBProject As Func(Of VBProject)
        Private ReadOnly _getOfficeApplication As Func(Of Object)
        Private ReadOnly _getApplicationInfo As Func(Of ApplicationInfo)
        Private ReadOnly _runCode As Func(Of String, Object)
        Private ReadOnly _runCodePreview As Func(Of String, Boolean, Boolean)
        Private ReadOnly _evaluateFormula As Func(Of String, Boolean, Boolean)

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        Public Sub New(
            getVBProject As Func(Of VBProject),
            getOfficeApplication As Func(Of Object),
            getApplicationInfo As Func(Of ApplicationInfo),
            runCode As Func(Of String, Object),
            runCodePreview As Func(Of String, Boolean, Boolean),
            evaluateFormula As Func(Of String, Boolean, Boolean))

            _getVBProject = getVBProject
            _getOfficeApplication = getOfficeApplication
            _getApplicationInfo = getApplicationInfo
            _runCode = runCode
            _runCodePreview = runCodePreview
            _evaluateFormula = evaluateFormula
        End Sub

#Region "辅助方法 - JSON修复"

        ''' <summary>
        ''' 智能修复 JSON 格式
        ''' </summary>
        Private Function FixJsonFormat(jsonStr As String) As String
            Try
                Dim fixedJson As String = jsonStr

                ' 1. 修复未加引号的属性名（在 { 或 , 后面）
                fixedJson = Regex.Replace(fixedJson, "([{,])\s*(\w+)\s*:", "$1""$2"":")

                ' 2. 修复转义引号问题
                fixedJson = fixedJson.Replace("\""", """")

                ' 3. 修复常见的格式问题
                fixedJson = FixMissingQuotes(fixedJson)

                ' 4. 尝试平衡大括号
                fixedJson = BalanceBraces(fixedJson)

                Return fixedJson
            Catch ex As Exception
                Debug.WriteLine($"FixJsonFormat 出错: {ex.Message}")
                Return jsonStr
            End Try
        End Function

        ''' <summary>
        ''' 修复缺失的引号
        ''' </summary>
        Private Function FixMissingQuotes(jsonStr As String) As String
            Dim result As New Text.StringBuilder()
            Dim inString As Boolean = False
            Dim escapeNext As Boolean = False
            Dim i As Integer = 0

            While i < jsonStr.Length
                Dim c As Char = jsonStr(i)

                If escapeNext Then
                    result.Append(c)
                    escapeNext = False
                ElseIf c = "\""" Then
                    inString = Not inString
                    result.Append(c)
                ElseIf c = "\" AndAlso inString Then
                    result.Append(c)
                    escapeNext = True
                ElseIf (c = "," OrElse c = "}") AndAlso Not inString Then
                    Dim lastQuoteIndex As Integer = result.ToString().LastIndexOf(""""c)
                    If lastQuoteIndex >= 0 Then
                        Dim betweenQuotesAndCurrent As String = result.ToString().Substring(lastQuoteIndex + 1)
                        If Not betweenQuotesAndCurrent.Contains(""""c) AndAlso betweenQuotesAndCurrent.Contains(":") Then
                            result.Append(""""c)
                        End If
                    End If
                    result.Append(c)
                Else
                    result.Append(c)
                End If

                i += 1
            End While

            Return result.ToString()
        End Function

        ''' <summary>
        ''' 平衡大括号
        ''' </summary>
        Private Function BalanceBraces(jsonStr As String) As String
            Dim openBraces As Integer = 0
            Dim closeBraces As Integer = 0

            For Each c As Char In jsonStr
                If c = "{" Then openBraces += 1
                If c = "}" Then closeBraces += 1
            Next

            Dim result As String = jsonStr.Trim()

            While openBraces > closeBraces
                result &= "}"
                closeBraces += 1
            End While

            Return result
        End Function

#End Region

#Region "代码执行入口"

        ''' <summary>
        ''' 根据语言类型执行代码
        ''' </summary>
        Public Sub ExecuteCode(code As String, language As String, preview As Boolean)
            Dim result = ExecuteCodeWithToolResult(code, language, preview)
            If result IsNot Nothing AndAlso Not result.Success Then
                GlobalStatusStrip.ShowWarning(If(result.UserMessage, result.Message))
            End If
        End Sub

        Public Function ExecuteCodeWithToolResult(code As String, language As String, preview As Boolean) As Agent.ToolResult
            Dim lowerLang As String = If(language, "").ToLower().Trim()

            If String.IsNullOrEmpty(lowerLang) OrElse lowerLang = "plaintext" OrElse lowerLang = "text" Then
                Dim trimmedCode = If(code, "").Trim()
                If trimmedCode.StartsWith("{") AndAlso trimmedCode.EndsWith("}") Then
                    Try
                        Dim testJson = Newtonsoft.Json.Linq.JObject.Parse(trimmedCode)
                        If testJson("command") IsNot Nothing OrElse testJson("commands") IsNot Nothing Then
                            lowerLang = "json"
                        End If
                    Catch
                    End Try
                End If
            End If

            If lowerLang.Contains("json") Then
                Return ExecuteJsonCommandWithToolResult(code, preview)
            End If

            Dim toolId As String = "ExecuteCode"
            Dim ok As Boolean
            If lowerLang.Contains("vbnet") OrElse lowerLang.Contains("vbscript") OrElse lowerLang.Contains("vba") Then
                toolId = "ExecuteVBA"
                ok = ExecuteVBACode(code, preview)
            ElseIf lowerLang.Contains("js") OrElse lowerLang.Contains("javascript") Then
                toolId = "ExecuteJavaScript"
                ok = ExecuteJavaScript(code, preview)
            ElseIf lowerLang.Contains("excel") OrElse lowerLang.Contains("formula") OrElse lowerLang.Contains("function") Then
                toolId = "ExecuteExcelFormula"
                ok = ExecuteExcelFormula(code, preview)
            ElseIf IsTextOnlyLanguage(lowerLang) Then
                Return Agent.ToolResult.Failed(toolId,
                                               $"文本类型 {language} 不可执行",
                                               errorCode:=ExceptionClassifier.CodeArgument,
                                               userMessage:="当前内容是文本，不是可执行命令",
                                               recoverable:=False)
            Else
                Return Agent.ToolResult.Failed(toolId,
                                               "不支持的语言类型: " & language,
                                               errorCode:=ExceptionClassifier.CodeArgument,
                                               userMessage:="不支持的语言类型: " & language,
                                               recoverable:=False)
            End If

            If CancelledByUser() Then
                Return Agent.ToolResult.PreviewCancelled(toolId)
            End If

            ' 脚本路径没有宿主快照：成功只说明 COM 调用没抛异常，失败也不能假定文档未被动过。
            Dim observation = New With {
                .kind = "code_execution",
                .summary = If(ok, $"{toolId} 执行完成（无快照校验）", $"{toolId} 执行失败"),
                .changed = ok,
                .targetRefs = New String() {"Office:ActiveDocument"},
                .warnings = New String() {"no_host_observer"}
            }
            If ok Then Return Agent.ToolResult.Succeed(toolId, observation.summary, observation:=observation)
            ' VBA/JS 可能在抛异常前已经改了文档，宿主既没有观察也没有回滚手段；
            ' recoverable=False 掐断自动重放，避免在已改过的文档上再跑一遍。
            Return Agent.ToolResult.Failed(toolId,
                                           observation.summary,
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:="脚本执行失败，可能已部分修改文档且无法自动撤销，请检查文档后再继续",
                                           recoverable:=False,
                                           observation:=observation)
        End Function

        ''' <summary>
        ''' JSON命令执行委托（由子类设置）
        ''' </summary>
        Public Property JsonCommandExecutorWithResult As Func(Of String, Boolean, Agent.ToolResult) = Nothing

        ''' <summary>
        ''' 宿主预览被用户拒绝时，Boolean 接口返回 True 表示“正常取消、不是错误”，
        ''' 但 ToolResult 合同不能据此回执成功。取消标记由宿主控件持有，这里只负责读取并清一次。
        ''' </summary>
        Public Property PreviewCancelledReader As Func(Of Boolean) = Nothing

        Private Function CancelledByUser() As Boolean
            Return PreviewCancelledReader IsNot Nothing AndAlso PreviewCancelledReader.Invoke()
        End Function

        Public Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
            Debug.WriteLine($"[CodeExecutionService] ExecuteJsonCommand (tool backend) preview={preview}")
            Debug.WriteLine($"[CodeExecutionService] JsonCommandExecutorWithResult set: {JsonCommandExecutorWithResult IsNot Nothing}")
            
            If JsonCommandExecutorWithResult IsNot Nothing Then
                Try
                    Dim currentJsonCode As String = jsonCode
                    Dim parseSuccess As Boolean = False

                    ' 首先尝试验证JSON是否有效
                    Try
                        Newtonsoft.Json.Linq.JObject.Parse(currentJsonCode)
                        parseSuccess = True
                    Catch
                        parseSuccess = False
                    End Try

                    ' 如果无效，尝试智能修复
                    If Not parseSuccess Then
                        Debug.WriteLine("[CodeExecutionService] JSON格式无效，尝试智能修复")
                        currentJsonCode = FixJsonFormat(jsonCode)
                        Debug.WriteLine($"[CodeExecutionService] 格式修正提示已生成，长度: {currentJsonCode.Length}")

                        ' 再次验证
                        Try
                            Newtonsoft.Json.Linq.JObject.Parse(currentJsonCode)
                            parseSuccess = True
                            Debug.WriteLine("[CodeExecutionService] JSON智能修复成功")
                        Catch
                            parseSuccess = False
                            Debug.WriteLine("[CodeExecutionService] JSON智能修复后仍然无效")
                        End Try
                    End If

                    Dim result As Agent.ToolResult = Nothing
                    ' 面板"执行"进来的 JSON 命令不经过 AgentKernel，日志全打成 cid=-，无法把工具失败归到某一次执行。
                    ' 已有作用域（Agent 主链正在跑）时既不覆盖也不替它清理。
                    Dim ownsLogScope = AppLogger.CorrelationId = "-"
                    If ownsLogScope Then AppLogger.BeginScope()
                    Try
                        result = JsonCommandExecutorWithResult.Invoke(currentJsonCode, preview)
                    Finally
                        If ownsLogScope Then AppLogger.ClearScope()
                    End Try
                    Debug.WriteLine($"[CodeExecutionService] JSON命令执行结果: {If(result Is Nothing, "null", result.ToObserveSummary())}")
                    If result Is Nothing Then Return Agent.ToolResult.Failed("", "JSON命令执行器未返回结果")
                    Return result
                Catch ex As Exception
                    Debug.WriteLine($"[CodeExecutionService] JSON命令执行异常: {ex.Message}")
                    GlobalStatusStrip.ShowWarning($"JSON命令执行失败: {ex.Message}")
                    Return Agent.ToolResult.FromException("", ex)
                End Try
            Else
                Debug.WriteLine("[CodeExecutionService] JsonCommandExecutorWithResult 未设置!")
                GlobalStatusStrip.ShowWarning("当前应用不支持JSON命令执行，请使用VBA代码")
                Return Agent.ToolResult.Failed("", "当前应用不支持JSON命令执行，请使用VBA代码")
            End If
        End Function

#End Region

#Region "VBA 代码执行"

        ''' <summary>
        ''' 所有脚本落笔点共用的前置门：与 Agent 工具路径同一个 SafetyGate 裁决。
        ''' Deny 直接拦截；NeedConfirm 走用户确认。宿主 UI、JSON 命令、旧聊天面板都必须先过这里。
        ''' </summary>
        Public Shared Function ConfirmScriptExecution(code As String, source As String) As Boolean
            If ScriptSafetyBlocked(code, source) Then Return False

            Dim decision = Agent.Execution.SafetyGate.EvaluateScript(code)
            If decision.Action = Agent.Execution.SafetyAction.RequireApproval Then
                If MessageBox.Show($"{decision.Reason}{vbCrLf}{vbCrLf}是否继续执行？",
                                   "需要确认",
                                   MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Question) <> DialogResult.Yes Then
                    Debug.WriteLine($"[{source}] 用户取消了需确认的脚本操作")
                    Return False
                End If
            End If
            Return True
        End Function

        ''' <summary>
        ''' 调用方已有自己的预览/确认界面时使用：只拦危险操作，不重复弹窗确认。
        ''' </summary>
        Public Shared Function ScriptSafetyBlocked(code As String, source As String) As Boolean
            Dim decision = Agent.Execution.SafetyGate.EvaluateScript(code)
            If decision.Action = Agent.Execution.SafetyAction.Deny Then
                MessageBox.Show($"安全拦截: {decision.Reason}{vbCrLf}{vbCrLf}该代码包含危险操作，禁止执行。",
                                "安全拦截",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Debug.WriteLine($"[{source}] 安全检查拦截: {decision.Reason}")
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' 执行 VBA 代码（增强版 - 集成安全检查）
        ''' </summary>
        Public Function ExecuteVBACode(vbaCode As String, preview As Boolean) As Boolean
            Try
                If Not ConfirmScriptExecution(vbaCode, "CodeExecutionService") Then Return False

                If preview Then
                    If Not _runCodePreview(vbaCode, preview) Then
                        Return True
                    End If
                End If

                Dim vbProj As VBProject = _getVBProject()
                If vbProj Is Nothing Then
                    Return False
                End If

                Dim vbComp As VBComponent = Nothing
                Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

                Try
                    ' 创建临时模块
                    vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
                    vbComp.Name = tempModuleName

                    If ContainsProcedureDeclaration(vbaCode) Then
                        ' 代码已包含过程声明
                        vbComp.CodeModule.AddFromString(vbaCode)
                        Dim procName As String = FindFirstProcedureName(vbComp)
                        If Not String.IsNullOrEmpty(procName) Then
                            _runCode(tempModuleName & "." & procName)
                        Else
                            GlobalStatusStrip.ShowWarning("无法在代码中找到可执行的过程")
                        End If
                    Else
                        ' 包装为过程
                        Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                                   vbaCode & vbNewLine &
                                                   "End Sub"
                        vbComp.CodeModule.AddFromString(wrappedCode)
                        _runCode(tempModuleName & ".Auto_Run")
                    End If

                    Return True
                Catch ex As Exception
                    MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                Finally
                    ' 删除临时模块
                    Try
                        If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                            vbProj.VBComponents.Remove(vbComp)
                        End If
                    Catch
                    End Try
                End Try
            Catch ex As Runtime.InteropServices.COMException
                HandleVBAException(ex)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 检查代码是否包含过程声明
        ''' </summary>
        Public Function ContainsProcedureDeclaration(code As String) As Boolean
            Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        End Function

        ''' <summary>
        ''' 查找模块中的第一个过程名
        ''' </summary>
        Public Function FindFirstProcedureName(comp As VBComponent) As String
            Try
                Dim codeModule As CodeModule = comp.CodeModule
                Dim lineCount As Integer = codeModule.CountOfLines
                Dim line As Integer = 1

                While line <= lineCount
                    Dim procName As String = codeModule.ProcOfLine(line, vbext_ProcKind.vbext_pk_Proc)
                    If Not String.IsNullOrEmpty(procName) Then
                        Return procName
                    End If
                    line = codeModule.ProcStartLine(procName, vbext_ProcKind.vbext_pk_Proc) + codeModule.ProcCountLines(procName, vbext_ProcKind.vbext_pk_Proc)
                End While

                Return String.Empty
            Catch
                ' 使用正则表达式提取
                Dim code As String = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                Dim match As Match = Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)

                If match.Success AndAlso match.Groups.Count > 2 Then
                    Return match.Groups(2).Value
                End If

                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 处理 VBA 异常
        ''' </summary>
        Private Sub HandleVBAException(ex As Runtime.InteropServices.COMException)
            If ex.Message.Contains("程序访问不被信任") OrElse
               ex.Message.Contains("Programmatic access to Visual Basic Project is not trusted") Then
                ShowVBATrustMessage()
            Else
                MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        ''' <summary>
        ''' 显示 VBA 信任设置提示
        ''' </summary>
        Private Sub ShowVBATrustMessage()
            MessageBox.Show(
                "无法执行 VBA 代码，请按以下步骤设置：" & vbCrLf & vbCrLf &
                "1. 点击 '文件' -> '选项' -> '信任中心'" & vbCrLf &
                "2. 点击 '信任中心设置'" & vbCrLf &
                "3. 选择 '宏设置'" & vbCrLf &
                "4. 勾选 '信任对 VBA 项目对象模型的访问'",
                "需要设置信任中心权限",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        End Sub

#End Region

#Region "JavaScript 执行"

        ''' <summary>
        ''' 执行 JavaScript 代码
        ''' </summary>
        Public Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
            Try
                ' 脚本引擎会把完整 Application 对象暴露给脚本，必须先过与 Agent 路径同一道门
                If Not ConfirmScriptExecution(jsCode, "CodeExecutionService") Then Return False

                Dim appObject As Object = _getOfficeApplication()
                If appObject Is Nothing Then
                    GlobalStatusStrip.ShowWarning("无法获取Office应用程序对象")
                    Return False
                End If

                ' 检测是否是 Office JS API 风格
                Dim isOfficeJsApiStyle As Boolean = jsCode.Contains("getActiveWorksheet") OrElse
                                                    jsCode.Contains("getUsedRange") OrElse
                                                    jsCode.Contains("getValues") OrElse
                                                    jsCode.Contains("setValues")

                ' 创建脚本控制引擎
                Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
                scriptEngine.Language = "JScript"

                ' 检测是否是 WPS
                Dim isWPS As Boolean = False
                Try
                    Dim appName As String = appObject.Name
                    isWPS = appName.Contains("WPS")
                Catch
                End Try

                ' 将 Office 应用对象暴露给脚本环境
                scriptEngine.AddObject("app", appObject, True)

                ' 添加适配层代码
                Dim adapterCode As String = GetJavaScriptAdapterCode(isWPS)
                scriptEngine.ExecuteStatement(adapterCode)

                ' 构建执行代码
                Dim wrappedCode As String = WrapJavaScriptCode(jsCode, isOfficeJsApiStyle)

                ' 执行并获取结果
                Dim result As String = scriptEngine.Eval(wrappedCode)
                GlobalStatusStrip.ShowInfo(result)

                Return True
            Catch ex As Exception
                MessageBox.Show("执行JavaScript代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取 JavaScript 适配层代码
        ''' </summary>
        Private Function GetJavaScriptAdapterCode(isWPS As Boolean) As String
            Return $"
            var Office = {{
                isWPS: {isWPS.ToString().ToLower()},
                app: app,
                context: {{
                    workbook: {{
                        getActiveWorksheet: function() {{
                            return {{
                                sheet: app.ActiveSheet,
                                getUsedRange: function() {{
                                    var usedRange = this.sheet.UsedRange;
                                    return {{
                                        range: usedRange,
                                        getValues: function() {{
                                            var values = [];
                                            var rows = this.range.Rows.Count;
                                            var cols = this.range.Columns.Count;
                                            for(var i = 1; i <= rows; i++) {{
                                                var rowValues = [];
                                                for(var j = 1; j <= cols; j++) {{
                                                    var cellValue = this.range.Cells(i, j).Value;
                                                    rowValues.push(cellValue);
                                                }}
                                                values.push(rowValues);
                                            }}
                                            return values;
                                        }},
                                        setValues: function(values) {{
                                            if(!values || values.length === 0) return;
                                            for(var i = 0; i < values.length; i++) {{
                                                var row = values[i];
                                                for(var j = 0; j < row.length; j++) {{
                                                    try {{
                                                        this.range.Cells(i+1, j+1).Value = row[j];
                                                    }} catch(e) {{ }}
                                                }}
                                            }}
                                        }}
                                    }};
                                }}
                            }};
                        }}
                    }}
                }},
                log: function(message) {{ return '输出: ' + message; }}
            }};
            function executeOfficeJsApi(codeFunc) {{
                var workbook = Office.context.workbook;
                if(typeof codeFunc === 'function') {{
                    try {{ return codeFunc(workbook); }}
                    catch(e) {{ return 'Office JS API 执行错误: ' + e.message; }}
                }}
                return 'Invalid function';
            }}
            "
        End Function

        ''' <summary>
        ''' 包装 JavaScript 代码
        ''' </summary>
        Private Function WrapJavaScriptCode(jsCode As String, isOfficeJsApiStyle As Boolean) As String
            If isOfficeJsApiStyle Then
                Return $"
                try {{
                    var userFunc = function(workbook) {{ {jsCode} }};
                    executeOfficeJsApi(userFunc);
                    return 'Office JS API 代码执行成功';
                }} catch(e) {{ return 'Office JS API 执行错误: ' + e.message; }}
                "
            Else
                Return $"
                try {{ {jsCode} return '代码执行成功'; }}
                catch(e) {{ return '执行错误: ' + e.message; }}
                "
            End If
        End Function

#End Region

#Region "Excel 公式执行"

        ''' <summary>
        ''' 执行 Excel 公式
        ''' </summary>
        Public Function ExecuteExcelFormula(formulaCode As String, preview As Boolean) As Boolean
            Try
                Dim appInfo As ApplicationInfo = _getApplicationInfo()

                ' 去除等号前缀
                If formulaCode.StartsWith("=") Then
                    formulaCode = formulaCode.Substring(1)
                End If

                If appInfo.Type = OfficeApplicationType.Excel Then
                    Dim result As Boolean = _evaluateFormula(formulaCode, preview)
                    GlobalStatusStrip.ShowInfo("公式执行结果: " & result.ToString())
                    Return True
                Else
                    GlobalStatusStrip.ShowWarning("Excel公式执行仅支持Excel环境")
                    Return False
                End If
            Catch ex As Exception
                MessageBox.Show("执行Excel公式时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#Region "辅助方法"

        ''' <summary>
        ''' 判断是否为纯文本类型语言（不可执行）
        ''' </summary>
        Private Function IsTextOnlyLanguage(language As String) As Boolean
            Dim textOnlyLanguages As String() = {
                "markdown", "md", "text", "plaintext", "txt",
                "html", "xml", "css", "yaml", "yml", "ini", "conf",
                "log", "diff", "patch", "csv", "tsv"
            }
            Return textOnlyLanguages.Any(Function(t) language.Contains(t))
        End Function

#End Region

    End Class
