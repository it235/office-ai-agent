' ShareRibbon\Controls\BaseChatControl.vb
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
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChatControl
    Inherits UserControl

    'Protected WithEvents ChatBrowser As WebView2
    'Protected WithEvents SelectedContentFlowPanel As FlowLayoutPanel
    Protected selectedCellChecked As Boolean = False
    'Protected _currentMarkdownBuffer As New StringBuilder()
    'Protected allMarkdownBuffer As New StringBuilder()

    ' ai的历史回复
    Protected historyMessageData As New List(Of HistoryMessage)

    Protected loadingPictureBox As PictureBox

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PASTE As Integer = &H302
        If m.Msg = WM_PASTE Then
            ' 在此处理粘贴操作，比如：
            If Clipboard.ContainsText() Then
                Dim txt As String = Clipboard.GetText()

                'QuestionTextBox.Text &= txt ' 将粘贴内容直接写入当前光标位置
            End If
            ' 不把消息传递给基类，从而拦截后续处理  
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Protected Async Function InitializeWebView2() As Task
        Try
            ' 自定义用户数据目录
            Dim userDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyAppWebView2Cache")

            ' 确保目录存在
            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            ' 释放资源文件到本地
            Dim wwwRoot As String = ResourceExtractor.ExtractResources()

            ' 配置 WebView2 的创建属性
            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                .UserDataFolder = userDataFolder
            }

            ' 初始化 WebView2
            Await ChatBrowser.EnsureCoreWebView2Async(Nothing)

            ' 确保 CoreWebView2 已初始化
            If ChatBrowser.CoreWebView2 IsNot Nothing Then

                ' 设置 WebView2 的安全选项
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True ' 开发时启用开发者工具

                ' 设置虚拟主机名映射到本地目录
                ChatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "officeai.local", ' 虚拟主机名
                    wwwRoot,          ' 本地文件夹路径
                    CoreWebView2HostResourceAccessKind.Allow  ' 允许访问
                )

                ' 替换模板中的 {wwwroot} 占位符
                Dim htmlContent As String = My.Resources.chat_template
                'htmlContent = htmlContent.Replace("{wwwroot}", wwwRoot.Replace("\", "/"))

                ' 修改HTML模板中的资源引用
                'Dim htmlContent As String = My.Resources.chat_template
                'htmlContent = htmlContent.Replace(
                '    "href=""css/",
                '    "href=""//officeai.local/css/"
                ').Replace(
                '    "src=""js/",
                '    "src=""//officeai.local/js/"
                ')

                ' 加载 HTML 模板
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)

                ' 配置 Marked 和代码高亮
                ConfigureMarked()
            Else
                MessageBox.Show("WebView2 初始化失败，CoreWebView2 不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    'Protected Async Function InitializeWebView2() As Task
    '    Try
    '        ' 检查 WebView2 是否已经初始化
    '        If ChatBrowser.CoreWebView2 IsNot Nothing Then
    '            Debug.WriteLine("WebView2 已经初始化，跳过创建过程")
    '            Return
    '        End If

    '        Debug.WriteLine("开始初始化 WebView2...")

    '        ' 自定义用户数据目录
    '        Dim userDataFolder As String = Path.Combine(
    '        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    '        "MyAppWebView2Cache"
    '    )

    '        ' 确保目录存在
    '        If Not Directory.Exists(userDataFolder) Then
    '            Directory.CreateDirectory(userDataFolder)
    '        End If

    '        ' 在UI线程上设置 CreationProperties
    '        If ChatBrowser.InvokeRequired Then
    '            Await ChatBrowser.Invoke(Sub()
    '                                         ' 创建新的 Environment 配置
    '                                         Dim envOptions = New CoreWebView2EnvironmentOptions()
    '                                         ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
    '                .UserDataFolder = userDataFolder
    '            }
    '                                     End Sub)
    '        Else
    '            ' 创建新的 Environment 配置
    '            Dim envOptions = New CoreWebView2EnvironmentOptions()
    '            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
    '            .UserDataFolder = userDataFolder
    '        }
    '        End If

    '        Debug.WriteLine("正在初始化 CoreWebView2...")

    '        ' 确保在UI线程上初始化 WebView2
    '        If ChatBrowser.InvokeRequired Then
    '            Await ChatBrowser.Invoke(Async Function()
    '                                         Await ChatBrowser.EnsureCoreWebView2Async(Nothing)
    '                                     End Function)
    '        Else
    '            Await ChatBrowser.EnsureCoreWebView2Async(Nothing)
    '        End If

    '        ' 确保 CoreWebView2 已初始化
    '        If ChatBrowser.CoreWebView2 IsNot Nothing Then
    '            Debug.WriteLine("CoreWebView2 初始化成功，正在加载模板...")

    '            ' 加载 HTML 模板
    '            If ChatBrowser.InvokeRequired Then
    '                ChatBrowser.Invoke(Sub() ChatBrowser.CoreWebView2.NavigateToString(My.Resources.chat_template))
    '            Else
    '                ChatBrowser.CoreWebView2.NavigateToString(My.Resources.chat_template)
    '            End If

    '            ' 配置 Marked 和代码高亮
    '            Await ConfigureMarked()

    '            Debug.WriteLine("模板加载完成")
    '        Else
    '            Throw New Exception("WebView2 初始化失败，CoreWebView2 不可用。")
    '        End If

    '    Catch ex As Exception
    '        Debug.WriteLine($"WebView2初始化失败: {ex.Message}")
    '        Debug.WriteLine($"堆栈跟踪: {ex.StackTrace}")
    '        Throw
    '    End Try
    'End Function

    Private Async Sub InjectScript(scriptContent As String)
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim escapedScript = JsonConvert.SerializeObject(scriptContent)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedScript})")
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法注入脚本。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Async Function ConfigureMarked() As Task
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim script = "
            marked.setOptions({
                highlight: function (code, lang) {
                    if (hljs.getLanguage(lang)) {
                        return hljs.highlight(lang, code).value;
                    } else {
                        return hljs.highlightAuto(code).value;
                    }
                }
            });
        "
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法配置 Marked。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Function

    ' 存储聊天HTML的文件路径
    Protected ReadOnly ChatHtmlFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
    )


    'Protected Async Sub InitializeWebView2()
    '    Try
    '        Dim userDataFolder As String = Path.Combine(
    '            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    '            "OfficeAiWebView2Cache"
    '        )

    '        Directory.CreateDirectory(userDataFolder)

    '        ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
    '            .UserDataFolder = userDataFolder
    '        }

    '        Await ChatBrowser.EnsureCoreWebView2Async()
    '        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived

    '        ' 加载HTML模板
    '        Await LoadLocalHtmlFile()
    '    Catch ex As Exception
    '        MessageBox.Show($"WebView2初始化失败: {ex.Message}")
    '    End Try
    'End Sub

    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim jsonDoc As JObject = JObject.Parse(e.WebMessageAsJson)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "checkedChange"
                    HandleCheckedChange(jsonDoc)
                Case "sendMessage"
                    HandleSendMessage(jsonDoc)
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case Else
                    Debug.WriteLine($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"处理消息出错: {ex.Message}")
        End Try
    End Sub

    Protected Overridable Sub HandleCheckedChange(jsonDoc As JObject)
        Dim prop As String = jsonDoc("property").ToString()
        Dim isChecked As Boolean = Boolean.Parse(jsonDoc("isChecked").ToString())
        If prop = "selectedCell" Then
            selectedCellChecked = isChecked
        End If
    End Sub

    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Dim question As String = jsonDoc("value").ToString()
        SendChatMessage(question)
    End Sub

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language)
    End Sub

    Protected MustOverride Function GetApplication() As Object
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCode(vbaCode As String)
    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法
    Private Sub ExecuteCode(code As String, language As String)
        ' 根据语言类型执行不同的操作
        Select Case language.ToLower()
            Case "vba", "vb", "vbscript", "language-vba", "language-vbscript", "language-vba hljs language-vbscript", "vba hljs language-vbscript"
                ' 执行 VBA 代码
                ExecuteVBACode(code)
            Case Else
                'MessageBox.Show("不支持的语言类型: " & language)
                GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
        End Select
    End Sub


    ' 执行前端传来的 VBA 代码片段
    Private Sub ExecuteVBACode(vbaCode As String)
        ' 获取 VBA 项目
        Dim vbProj As VBProject = GetVBProject()

        ' 添加空值检查
        If vbProj Is Nothing Then
            Return
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' 创建临时模块
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' 检查代码是否已包含 Sub/Function 声明
            If ContainsProcedureDeclaration(vbaCode) Then
                ' 代码已包含过程声明，直接添加
                vbComp.CodeModule.AddFromString(vbaCode)

                ' 查找第一个过程名并执行
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    RunCode(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("无法在代码中找到可执行的过程")
                    GlobalStatusStrip.ShowWarning("无法在代码中找到可执行的过程")
                End If
            Else
                ' 代码不包含过程声明，将其包装在 Auto_Run 过程中
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' 执行 Auto_Run 过程
                RunCode(tempModuleName & ".Auto_Run")
            End If

        Catch ex As Exception
            MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 无论成功还是失败，都删除临时模块
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' 忽略清理错误
            End Try
        End Try
    End Sub


    ' 检查代码是否包含过程声明
    Private Function ContainsProcedureDeclaration(code As String) As Boolean
        ' 使用简单的正则表达式检查是否包含 Sub 或 Function 声明
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function


    ' 查找模块中的第一个过程名
    Private Function FindFirstProcedureName(comp As VBComponent) As String
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
            ' 如果出错，尝试使用正则表达式从代码中提取
            Dim code As String = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            Dim match As Match = Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)

            If match.Success AndAlso match.Groups.Count > 2 Then
                Return match.Groups(2).Value
            End If

            Return String.Empty
        End Try
    End Function

    Public Async Function Send(question As String) As Task
        Dim apiUrl As String = ConfigSettings.ApiUrl
        Dim apiKey As String = ConfigSettings.ApiKey

        If String.IsNullOrWhiteSpace(apiKey) Then
            GlobalStatusStrip.ShowWarning("请先配置大模型ApiKey！")
            Return
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            GlobalStatusStrip.ShowWarning("请先配置大模型Api！")
            Return
        End If

        If String.IsNullOrWhiteSpace(question) Then
            GlobalStatusStrip.ShowWarning("请输入问题！")
            Return
        End If

        Dim uuid As String = Guid.NewGuid().ToString()

        Try
            Dim requestBody As String = CreateRequestBody(question)
            Await SendHttpRequestStream(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)
            Await SaveFullWebPageAsync2()
        Catch ex As Exception
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try

    End Function


    Private Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")

        ' 构建 messages 数组
        Dim messages As New List(Of String)()

        ' 添加 system 消息
        Dim systemMessage = historyMessageData.FirstOrDefault(Function(m) m.role = "system")
        If systemMessage IsNot Nothing Then
            historyMessageData.Remove(systemMessage)
        End If
        systemMessage = New HistoryMessage() With {
            .role = "system",
            .content = ConfigSettings.propmtContent
        }
        historyMessageData.Insert(0, systemMessage)

        Dim q = New HistoryMessage() With {
                .role = "user",
                .content = result
            }
        historyMessageData.Add(q)

        ' 添加历史消息
        For Each message In historyMessageData
            messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{message.content}""}}")
        Next

        ' 构建 JSON 请求体
        Dim messagesJson = String.Join(",", messages)
        Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""messages"": [{messagesJson}], ""stream"": true}}"

        Return requestBody
        ' 使用从 ConfigSettings 中获取的模型名称
        'Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""system"", ""content"": """ & ConfigSettings.propmtContent & """},{""role"": ""user"", ""content"": """ & result & """}],""stream"":true}"
    End Function


    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String) As Task
        Try

            ' 强制使用 TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = Timeout.InfiniteTimeSpan

                ' 准备请求 ---
                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                ' 打印请求日志 ---
                Debug.WriteLine("[HTTP] 开始发送流式请求...")
                Debug.WriteLine($"[HTTP] Request Body: {requestBody}")

                ' 组装ai答复头部
                Dim uuid As String = Guid.NewGuid().ToString()

                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                ' 发送请求 ---
                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()
                    Debug.WriteLine($"[HTTP] 响应状态码: {response.StatusCode}")

                    Dim js As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{uuid}');"
                    If ChatBrowser.InvokeRequired Then
                        ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(js))
                    Else
                        Await ChatBrowser.ExecuteScriptAsync(js)
                    End If

                    ' 处理流 ---
                    Dim stringBuilder As New StringBuilder()
                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim buffer(10230) As Char
                            Dim readCount As Integer
                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunk As String = New String(buffer, 0, readCount)
                                ' 如果chunk不是以data开头，则跳过
                                'If Not chunk.StartsWith("data:") Then Continue Do
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)
                                'Debug.WriteLine($"[Stream] 接收到块:{stringBuilder}")
                                ' 判断stringBuilder是否以'}'结尾，如果是则解析
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), uuid)
                                    stringBuilder.Clear()
                                End If

                                'If Not line.StartsWith("{") OrElse Not line.EndsWith("}") Then
                                '    _currentMarkdownBuffer.Append(line)
                                '    Continue For
                                'End If

                            Loop
                            Debug.WriteLine("[Stream] 流接收完成")
                            Await ExecuteJavaScriptAsyncJS($"processStreamComplete('{uuid}');")
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 请求过程中出错: {ex.ToString()}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 记录全局上下文中，方便后续使用
            Dim answer = New HistoryMessage() With {
                .role = "assistant",
                .content = allMarkdownBuffer.ToString()
            }
            historyMessageData.Add(answer)
            allMarkdownBuffer.Clear()
        End Try
    End Function



    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' 启用表格、代码块等扩展
    .Build()                        ' 构建不可变管道

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()

    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String)
        Try
            'Dim lines As String() = rawChunk.Split({"data:"}, StringSplitOptions.RemoveEmptyEntries)
            Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                line = line.Trim()
                If line = "[DONE]" Then
                    FlushBuffer("content", uuid) ' 最后刷新缓冲区
                    Return
                End If
                If line = "" Then
                    Continue For
                End If

                Debug.Print(line)
                Dim jsonObj As JObject = JObject.Parse(line)
                Dim reasoning_content As String = jsonObj("choices")(0)("delta")("reasoning_content")?.ToString()
                If Not String.IsNullOrEmpty(reasoning_content) Then
                    _currentMarkdownBuffer.Append(reasoning_content)
                    ' 检查是否到达代码块自然分割点（例如换行符）
                    'If reasoning_content.Contains(vbLf) OrElse reasoning_content.TrimEnd().EndsWith("`") Then
                    FlushBuffer("reasoning", uuid)
                    'End If
                End If

                Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()

                If Not String.IsNullOrEmpty(content) Then
                    _currentMarkdownBuffer.Append(content)
                    ' 检查是否到达代码块自然分割点（例如换行符）
                    'If content.Contains(vbLf) OrElse content.TrimEnd().EndsWith("`") Then
                    FlushBuffer("content", uuid)
                    'End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 数据处理失败: {ex.Message}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Async Sub FlushBuffer(contentType As String, uuid As String)
        If _currentMarkdownBuffer.Length = 0 Then Return

        Dim escapedContent = HttpUtility.JavaScriptStringEncode(_currentMarkdownBuffer.ToString())
        _currentMarkdownBuffer.Clear()
        Dim js As String
        If contentType = "reasoning" Then
            js = $"appendReasoning('{uuid}','{escapedContent}');"
        Else
            js = $"appendRenderer('{uuid}','{escapedContent}');"
            allMarkdownBuffer.Append(escapedContent)
        End If

        Await ExecuteJavaScriptAsyncJS(js)
    End Sub


    ' 执行js脚本的异步方法
    Private Async Function ExecuteJavaScriptAsyncJS(js As String) As Task
        If ChatBrowser.InvokeRequired Then
            ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(js))
        Else
            Await ChatBrowser.ExecuteScriptAsync(js)
        End If
    End Function

    Private Function DecodeBase64(base64 As String) As String
        Dim bytes As Byte() = System.Convert.FromBase64String(base64)
        Return System.Text.Encoding.UTF8.GetString(bytes)
    End Function

    Private Function EscapeJavaScriptString(input As String) As String
        Return input _
        .Replace("\", "\\") _
        .Replace("'", "\'") _
        .Replace(vbCr, "") _
        .Replace(vbLf, "\n") _
        .Replace("</script>", "<\/script>")  ' 避免脚本注入
    End Function



    ' 共用的HTTP请求方法
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            MessageBox.Show($"请求失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 加载本地HTML文件
    Public Async Function LoadLocalHtmlFile() As Task
        Try
            ' 检查HTML文件是否存在
            Dim htmlFilePath As String = ChatHtmlFilePath
            If File.Exists(htmlFilePath) Then

                Await Task.Run(Sub()
                                   Dim htmlContent As String = File.ReadAllText(htmlFilePath, System.Text.Encoding.UTF8)
                                   htmlContent = htmlContent.TrimStart("""").TrimEnd("""")
                                   ' 直接导航到本地HTML文件
                                   If ChatBrowser.InvokeRequired Then
                                       ChatBrowser.Invoke(Sub() ChatBrowser.CoreWebView2.NavigateToString(htmlContent))
                                   Else
                                       ChatBrowser.CoreWebView2.NavigateToString(htmlContent)
                                   End If
                               End Sub)

            End If
        Catch ex As Exception
            Debug.WriteLine($"加载本地HTML文件时出错：{ex.Message}")
        End Try
    End Function

    Public Async Function SaveFullWebPageAsync2() As Task
        Try
            ' 1. 创建目录（同步操作，无需异步）

            Dim dir = Path.GetDirectoryName(ChatHtmlFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' 2. 获取 HTML（异步无阻塞）
            Dim htmlContent As String = Await GetFullHtmlContentAsync()

            ' 3. 保存文件（异步后台线程）
            Await Task.Run(Sub()
                               Dim fullHtml As String = "<!DOCTYPE html>" & Environment.NewLine & htmlContent
                               File.WriteAllText(
                $"{ChatHtmlFilePath}",
                HttpUtility.HtmlDecode(fullHtml),
                System.Text.Encoding.UTF8
            )
                           End Sub)

            Debug.WriteLine("保存成功")
        Catch ex As Exception
            Debug.WriteLine($"保存失败: {ex.Message}")
        End Try
    End Function

    Private Async Function GetFullHtmlContentAsync() As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' 强制切换到 WebView2 的 UI 线程操作
        ChatBrowser.BeginInvoke(Async Sub()
                                    Try
                                        Await EnsureWebView2InitializedAsync()

                                        Dim js As String = "
                (function(){
                    const serializer = new XMLSerializer();
                    return serializer.serializeToString(document.documentElement);
                })();"

                                        Dim rawResult As String = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(js)
                                        Dim decodedHtml As String = UnescapeHtmlContent(rawResult)
                                        decodedHtml = decodedHtml.TrimStart("""").TrimEnd("""")

                                        ' 移除 <script> 标签及其内容
                                        Dim scriptPattern As String = "<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>"
                                        decodedHtml = Regex.Replace(decodedHtml, scriptPattern, String.Empty, RegexOptions.IgnoreCase)

                                        tcs.SetResult(decodedHtml)
                                    Catch ex As Exception
                                        tcs.SetException(ex)
                                    End Try
                                End Sub)

        Return Await tcs.Task
    End Function

    Private Async Function EnsureWebView2InitializedAsync() As Task
        If ChatBrowser.CoreWebView2 Is Nothing Then
            Await ChatBrowser.EnsureCoreWebView2Async()
        End If
    End Function

    Private Function UnescapeHtmlContent(htmlContent As String) As String
        ' 处理转义字符（直接从 JSON 字符串中提取）
        Return System.Text.RegularExpressions.Regex.Unescape(
        htmlContent
    )
    End Function

    ' 提示词配置（每次仅可使用1个）
    Public Class HistoryMessage
        Public Property role As String
        Public Property content As String
    End Class

    ' 注入辅助脚本
    Protected Sub InitializeWebView2Script()
        ' 设置 Web 消息处理器
        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived

        ' 检查本地HTML文件是否存在
        Dim htmlFilePath As String = ChatHtmlFilePath
        If File.Exists(htmlFilePath) Then
            ' 加载本地HTML文件
            LoadLocalHtmlFile()
        End If

        ' 注入辅助脚本
        Dim script As String = "
        window.vsto = {
            executeCode: function(code, language) {
                window.chrome.webview.postMessage({
                    type: 'executeCode',
                    code: code,
                    language: language
                });
                return true;
            },
            checkedChange: function(thisProperty,checked) {
                return window.chrome.webview.postMessage({
                    type: 'checkedChange',
                    isChecked: checked,
                    property: thisProperty
                });
            },
            sendMessage: function(question) {
                return window.chrome.webview.postMessage({
                    type: 'sendMessage',
                    value: question
                });
            }
        };
    "
        ChatBrowser.ExecuteScriptAsync(script)
    End Sub

    ' 选中内容发送到聊天区
    Protected Async Sub AddSelectedContentItem(sheetName As String, address As String)
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub


    Protected Shared Sub VBAxceptionHandle(ex As Runtime.InteropServices.COMException)
        ' 处理信任中心权限问题
        If ex.Message.Contains("程序访问不被信任") OrElse
       ex.Message.Contains("Programmatic access to Visual Basic Project is not trusted") Then
            VBATrustShowBox()
        Else
            MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Shared Sub VBATrustShowBox()
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

End Class