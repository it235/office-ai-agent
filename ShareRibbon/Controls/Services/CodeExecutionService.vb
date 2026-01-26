' ShareRibbon\Controls\Services\CodeExecutionService.vb
' 代码执行服务：VBA、JavaScript、Excel公式执行

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

#Region "代码执行入口"

        ''' <summary>
        ''' 根据语言类型执行代码
        ''' </summary>
        Public Sub ExecuteCode(code As String, language As String, preview As Boolean)
            Dim lowerLang As String = If(language, "").ToLower().Trim()
            
            ' 调试日志
            Debug.WriteLine($"[CodeExecutionService] 执行代码, 语言: '{language}' (规范化: '{lowerLang}')")
            Debug.WriteLine($"[CodeExecutionService] 代码前100字符: {If(code.Length > 100, code.Substring(0, 100), code)}...")
            
            ' 自动检测JSON：如果代码看起来像JSON命令格式
            If String.IsNullOrEmpty(lowerLang) OrElse lowerLang = "plaintext" OrElse lowerLang = "text" Then
                Dim trimmedCode = code.Trim()
                If trimmedCode.StartsWith("{") AndAlso trimmedCode.EndsWith("}") Then
                    Try
                        Dim testJson = Newtonsoft.Json.Linq.JObject.Parse(trimmedCode)
                        If testJson("command") IsNot Nothing Then
                            Debug.WriteLine("[CodeExecutionService] 自动检测为JSON命令格式")
                            lowerLang = "json"
                        End If
                    Catch
                        ' 不是有效的JSON命令
                    End Try
                End If
            End If

            If lowerLang.Contains("json") Then
                Debug.WriteLine("[CodeExecutionService] 路由到JSON命令执行器")
                ExecuteJsonCommand(code, preview)
            ElseIf lowerLang.Contains("vbnet") OrElse lowerLang.Contains("vbscript") OrElse lowerLang.Contains("vba") Then
                Debug.WriteLine("[CodeExecutionService] 路由到VBA执行器")
                ExecuteVBACode(code, preview)
            ElseIf lowerLang.Contains("js") OrElse lowerLang.Contains("javascript") Then
                Debug.WriteLine("[CodeExecutionService] 路由到JavaScript执行器")
                ExecuteJavaScript(code, preview)
            ElseIf lowerLang.Contains("excel") OrElse lowerLang.Contains("formula") OrElse lowerLang.Contains("function") Then
                Debug.WriteLine("[CodeExecutionService] 路由到Excel公式执行器")
                ExecuteExcelFormula(code, preview)
            Else
                Debug.WriteLine($"[CodeExecutionService] 不支持的语言类型: '{language}'")
                GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
            End If
        End Sub

        ''' <summary>
        ''' JSON命令执行委托（由子类设置）
        ''' </summary>
        Public Property JsonCommandExecutor As Func(Of String, Boolean, Boolean) = Nothing

        ''' <summary>
        ''' 执行JSON命令
        ''' </summary>
        Public Sub ExecuteJsonCommand(jsonCode As String, preview As Boolean)
            Debug.WriteLine($"[CodeExecutionService] ExecuteJsonCommand 被调用, preview={preview}")
            Debug.WriteLine($"[CodeExecutionService] JsonCommandExecutor 是否已设置: {JsonCommandExecutor IsNot Nothing}")
            
            If JsonCommandExecutor IsNot Nothing Then
                Try
                    Dim result = JsonCommandExecutor.Invoke(jsonCode, preview)
                    Debug.WriteLine($"[CodeExecutionService] JSON命令执行结果: {result}")
                Catch ex As Exception
                    Debug.WriteLine($"[CodeExecutionService] JSON命令执行异常: {ex.Message}")
                    GlobalStatusStrip.ShowWarning($"JSON命令执行失败: {ex.Message}")
                End Try
            Else
                Debug.WriteLine("[CodeExecutionService] JsonCommandExecutor 未设置!")
                GlobalStatusStrip.ShowWarning("当前应用不支持JSON命令执行，请使用VBA代码")
            End If
        End Sub

#End Region

#Region "VBA 代码执行"

        ''' <summary>
        ''' 执行 VBA 代码
        ''' </summary>
        Public Function ExecuteVBACode(vbaCode As String, preview As Boolean) As Boolean
            Try
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

    End Class
