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
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Markdig
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChatControl
    Inherits UserControl

    'settings
    Protected topicRandomness As Double
    Protected contextLimit As Integer
    Protected selectedCellChecked As Boolean = False
    Protected settingsScrollChecked As Boolean = False

    Protected stopReaderStream As Boolean = False


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

                ' 加载 HTML 模板
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)

                ' 配置 Marked 和代码高亮
                ConfigureMarked()
                ' 添加导航完成事件处理，确保在页面加载完成后初始化设置
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted

            Else
                MessageBox.Show("WebView2 初始化失败，CoreWebView2 不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
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


    ' 动态ChatHtmlFilePath属性
    Protected ReadOnly Property ChatHtmlFilePath As String
        Get
            ' 如果已经生成过文件路径，直接返回缓存的路径
            If Not String.IsNullOrEmpty(_chatHtmlFilePath) Then
                Return _chatHtmlFilePath
            End If

            Dim baseDir As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder
        )

            Dim fileName As String
            If Not String.IsNullOrEmpty(firstQuestion) Then
                ' 简单地取前10个字符
                Dim questionPrefix As String = GetFirst10Characters(firstQuestion)
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}_{questionPrefix}.html"
            Else
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            End If

            _chatHtmlFilePath = Path.Combine(baseDir, fileName)
            Return _chatHtmlFilePath
        End Get
    End Property
    Private Function GetFirst10Characters(text As String) As String
        If String.IsNullOrEmpty(text) Then Return String.Empty

        ' 取前20个字符，如果不足20个则取全部
        Dim result As String = If(text.Length > 20, text.Substring(0, 20), text)

        ' 移除文件名中不允许的字符，替换为下划线
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        For Each invalidChar In invalidChars
            result = result.Replace(invalidChar, "_"c)
        Next

        ' 移除一些可能导致问题的字符
        result = result.Replace(" ", "_")  ' 空格替换为下划线
        result = result.Replace(".", "_")  ' 点号替换为下划线
        result = result.Replace(",", "_")  ' 逗号替换为下划线
        result = result.Replace(":", "_")  ' 冒号替换为下划线
        result = result.Replace("?", "_")  ' 问号替换为下划线
        result = result.Replace("!", "_")  ' 感叹号替换为下划线

        Return result
    End Function
    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If e.IsSuccess Then
            Try
                If ChatBrowser.InvokeRequired Then
                    ' 使用同步的 Invoke 而不是异步的
                    ChatBrowser.Invoke(Sub()
                                           Task.Delay(100).Wait() ' 同步等待
                                           InitializeSettings()
                                           InitializeMcpSettings() ' 添加MCP初始化

                                           ' 直接在UI线程移除事件处理器
                                           If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                                               RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                                           End If
                                       End Sub)
                Else
                    Task.Delay(100).Wait() ' 同步等待
                    InitializeSettings()
                    InitializeMcpSettings() ' 添加MCP初始化

                    ' 直接在UI线程移除事件处理器
                    If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                        RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        End If
    End Sub

    Protected Sub InitializeSettings()
        Try
            ' 加载设置
            Dim chatSettings As New ChatSettings(GetApplication())
            selectedCellChecked = ChatSettings.selectedCellChecked
            contextLimit = ChatSettings.contextLimit
            topicRandomness = ChatSettings.topicRandomness
            settingsScrollChecked = ChatSettings.settingsScrollChecked

            ' 将设置发送到前端
            Dim js As String = $"
            document.getElementById('topic-randomness').value = '{ChatSettings.topicRandomness}';
            document.getElementById('topic-randomness-value').textContent = '{ChatSettings.topicRandomness}';
            document.getElementById('context-limit').value = '{ChatSettings.contextLimit}';
            document.getElementById('context-limit-value').textContent = '{ChatSettings.contextLimit}';
            document.getElementById('settings-scroll-checked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('settings-selected-cell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
            document.getElementById('settings-executecode-preview').checked = {chatSettings.executecodePreviewChecked.ToString().ToLower()};
            
            var selectElement = document.getElementById('chatMode');
            if (selectElement) {{
                selectElement.value = '{ChatSettings.chatMode}';
            }}
            
            // 同步到主界面的checkbox
            document.getElementById('scrollChecked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('selectedCell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
        "
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"初始化设置失败: {ex.Message}")
        End Try
    End Sub

    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim jsonDoc As JObject = JObject.Parse(e.WebMessageAsJson)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "checkedChange"
                    HandleCheckedChange(jsonDoc)
                Case "sendMessage"
                    HandleSendMessage(jsonDoc)
                Case "stopMessage"
                    stopReaderStream = True
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case "saveSettings"
                    Debug.Print("保存设置")
                    HandleSaveSettings(jsonDoc)
                Case "getHistoryFiles"
                    HandleGetHistoryFiles()
                Case "openHistoryFile"
                    HandleOpenHistoryFile(jsonDoc)
                Case "getMcpConnections"
                    HandleGetMcpConnections()
                Case "saveMcpSettings"
                    HandleSaveMcpSettings(jsonDoc)
                Case "clearContext"
                    ClearChatContext()
                Case Else
                    Debug.WriteLine($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"处理消息出错: {ex.Message}")
        End Try
    End Sub

    Private Sub ClearChatContext()
        historyMessageData.Clear()
        Debug.WriteLine("已清空聊天记忆（上下文）")
    End Sub

    ' 处理获取MCP连接列表请求
    Protected Sub HandleGetMcpConnections()
        Try
            ' 获取所有可用的MCP连接
            Dim connections = MCPConnectionManager.LoadConnections()

            ' 过滤出已启用的连接 - 使用IsActive而不是Enabled
            Dim enabledConnections = connections.Where(Function(c) c.IsActive).ToList()

            ' 获取已启用的MCP列表
            Dim chatSettings As New ChatSettings(GetApplication())
            Dim enabledMcpList = chatSettings.EnabledMcpList

            ' 将数据序列化为JSON
            Dim connectionsJson = JsonConvert.SerializeObject(enabledConnections)
            Dim enabledListJson = JsonConvert.SerializeObject(enabledMcpList)

            ' 获取当前配置的模型是否支持mcp
            Dim mcpSupported As Boolean = ConfigSettings.mcpable
            'Debug.WriteLine($"获取MCP连接列表，当前模型是否支持MCP {mcpSupported}")

            ' 发送到前端
            Dim js = $"renderMcpConnections({connectionsJson}, {enabledListJson},{mcpSupported.ToString().ToLower()});"
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"获取MCP连接列表失败: {ex.Message}")
        End Try
    End Sub

    ' 处理保存MCP设置请求
    Protected Sub HandleSaveMcpSettings(jsonDoc As JObject)
        Try
            ' 获取启用的MCP列表
            Dim enabledList As List(Of String) = jsonDoc("enabledList").ToObject(Of List(Of String))()

            ' 保存到设置
            Dim chatSettings As New ChatSettings(GetApplication())
            chatSettings.SaveEnabledMcpList(enabledList)

            GlobalStatusStrip.ShowInfo("MCP设置已保存")
        Catch ex As Exception
            Debug.WriteLine($"保存MCP设置失败: {ex.Message}")
            GlobalStatusStrip.ShowWarning("保存MCP设置失败")
        End Try
    End Sub

    ' 添加MCP初始化方法
    Protected Sub InitializeMcpSettings()
        Try
            ' 检查当前模型是否支持MCP
            Dim mcpSupported = False

            ' 从ConfigData中查找当前模型的mcpable属性
            For Each config In ConfigManager.ConfigData
                If config.selected Then
                    For Each model In config.model
                        If model.selected Then
                            mcpSupported = model.mcpable
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next

            ' 加载MCP连接和启用列表
            Dim connections = MCPConnectionManager.LoadConnections()
            ' 使用IsActive替代Enabled
            Dim enabledConnections = connections.Where(Function(c) c.IsActive).ToList()

            Dim chatSettings As New ChatSettings(GetApplication())
            Dim enabledMcpList = chatSettings.EnabledMcpList

            ' 序列化为JSON
            Dim connectionsJson = JsonConvert.SerializeObject(enabledConnections)
            Dim enabledListJson = JsonConvert.SerializeObject(enabledMcpList)

            ' 向前端传递信息
            Dim js = $"setMcpSupport({mcpSupported.ToString().ToLower()}, {connectionsJson}, {enabledListJson});"
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"初始化MCP设置失败: {ex.Message}")
        End Try
    End Sub
    ' 处理获取历史文件列表请求
    Protected Sub HandleGetHistoryFiles()
        Try
            ' 获取历史文件目录
            Dim historyDir As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder
        )

            Dim historyFiles As New List(Of Object)()

            If Directory.Exists(historyDir) Then
                ' 查找所有符合条件的HTML文件
                Dim files As String() = Directory.GetFiles(historyDir, "saved_chat_*.html")

                For Each filePath As String In files
                    Try
                        Dim fileInfo As New FileInfo(filePath)
                        historyFiles.Add(New With {
                        .fileName = fileInfo.Name,
                        .fullPath = fileInfo.FullName,
                        .size = fileInfo.Length,
                        .lastModified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    })
                    Catch ex As Exception
                        Debug.WriteLine($"处理文件信息时出错: {filePath} - {ex.Message}")
                    End Try
                Next
            End If

            ' 将文件列表序列化为JSON并发送到前端
            Dim jsonResult As String = JsonConvert.SerializeObject(historyFiles)
            Dim escapedJson As String = HttpUtility.JavaScriptStringEncode(jsonResult)

            Dim js As String = $"setHistoryFilesList({jsonResult});"
            ExecuteJavaScriptAsyncJS(js)

        Catch ex As Exception
            Debug.WriteLine($"获取历史文件列表时出错: {ex.Message}")
            ' 发送空列表到前端
            ExecuteJavaScriptAsyncJS("setHistoryFilesList([]);")
        End Try
    End Sub

    ' 处理打开历史文件请求
    Protected Sub HandleOpenHistoryFile(jsonDoc As JObject)
        Try
            Dim filePath As String = jsonDoc("filePath").ToString()

            If File.Exists(filePath) Then
                ' 使用默认浏览器打开HTML文件
                Process.Start(New ProcessStartInfo() With {
                .FileName = filePath,
                .UseShellExecute = True
            })

                GlobalStatusStrip.ShowInfo("已在浏览器中打开历史记录")
            Else
                GlobalStatusStrip.ShowWarning("历史记录文件不存在")
            End If

        Catch ex As Exception
            Debug.WriteLine($"打开历史文件时出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("打开历史记录失败: " & ex.Message)
        End Try
    End Sub

    Protected Overridable Sub HandleCheckedChange(jsonDoc As JObject)
        Dim prop As String = jsonDoc("property").ToString()
        Dim isChecked As Boolean = Boolean.Parse(jsonDoc("isChecked").ToString())
        If prop = "selectedCell" Then
            selectedCellChecked = isChecked
        End If
    End Sub

    Protected Overridable Sub HandleSaveSettings(jsonDoc As JObject)
        topicRandomness = jsonDoc("topicRandomness")
        contextLimit = jsonDoc("contextLimit")
        selectedCellChecked = jsonDoc("selectedCell")
        settingsScrollChecked = jsonDoc("settingsScroll")
        Dim chatMode As String = jsonDoc("chatMode")
        Dim executeCodePreview As Boolean = jsonDoc("executeCodePreview")
        Dim chatSettings As New ChatSettings(GetApplication())
        ' 保存设置到配置文件
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, executeCodePreview, chatMode)
    End Sub

    Public Class SendMessageReferenceContentItem
        Public Property id As String
        Public Property sheetName As String
        Public Property address As String
    End Class

    ' 定义文件内容解析结果的类
    Public Class FileContentResult
        Public Property FileName As String
        Public Property FileType As String  ' "Excel", "Word", "Text", 等
        Public Property ParsedContent As String  ' 格式化的内容字符串
        Public Property RawData As Object  ' 原始数据，可用于进一步处理
    End Class


    ' 添加存储第一个问题的变量
    Protected firstQuestion As String = String.Empty
    Protected isFirstMessage As Boolean = True
    Private _chatHtmlFilePath As String = String.Empty ' 缓存文件路径



    ' 在 HandleSendMessage 方法中添加文件内容解析逻辑
    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String
        Dim filePaths As List(Of String) = New List(Of String)()
        Dim selectedContents As List(Of SendMessageReferenceContentItem) = New List(Of SendMessageReferenceContentItem)()

        If messageValue.Type = JTokenType.Object Then
            ' New format with text, potentially filePaths, and selectedContent
            question = messageValue("text")?.ToString()

            If messageValue("filePaths") IsNot Nothing AndAlso messageValue("filePaths").Type = JTokenType.Array Then
                filePaths = messageValue("filePaths").ToObject(Of List(Of String))()
            End If

            ' 解析 selectedContent
            If messageValue("selectedContent") IsNot Nothing AndAlso messageValue("selectedContent").Type = JTokenType.Array Then
                Try
                    selectedContents = messageValue("selectedContent").ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If
        Else
            Debug.WriteLine("HandleSendMessage: Invalid message format for 'value'.")
            Return
        End If

        If String.IsNullOrEmpty(question) AndAlso
       (filePaths Is Nothing OrElse filePaths.Count = 0) AndAlso
       (selectedContents Is Nothing OrElse selectedContents.Count = 0) Then
            Debug.WriteLine("HandleSendMessage: Empty question, no files, and no selected content.")
            Return ' Nothing to send
        End If

        ' 保存第一个问题（仅保存一次）
        If isFirstMessage AndAlso Not String.IsNullOrEmpty(question) Then
            firstQuestion = question
            isFirstMessage = False
            ' 清空缓存的文件路径，强制重新生成
            _chatHtmlFilePath = String.Empty
            Debug.WriteLine($"保存第一个问题: {firstQuestion}")
            Debug.WriteLine($"将生成文件路径: {ChatHtmlFilePath}")
        End If

        ' --- 处理选中的内容 ---
        question = AppendCurrentSelectedContent("--- 我此次的问题：" & question & " ---")

        ' --- 文件内容解析逻辑 ---
        Dim fileContentBuilder As New StringBuilder()
        Dim parsedFiles As New List(Of FileContentResult)()

        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            question = question & " 用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---"

            fileContentBuilder.AppendLine(vbCrLf & "--- 以下是用户引用的其他文件内容 ---")

            ' 获取当前工作目录
            Dim currentWorkingDir As String = GetCurrentWorkingDirectory()
            If String.IsNullOrEmpty(currentWorkingDir) Then
                GlobalStatusStrip.ShowWarning("请保存当前文件，并且把引用文件和当前文件放在同一目录下后重试: ")
                Return
            End If

            For Each filePath As String In filePaths
                Try
                    ' 检查文件是否为绝对路径
                    Dim fullFilePath As String = filePath

                    ' 如果不是绝对路径，则尝试在当前工作目录下查找
                    If Not Path.IsPathRooted(filePath) AndAlso Not String.IsNullOrEmpty(currentWorkingDir) Then
                        fullFilePath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                        Debug.WriteLine($"尝试在工作目录查找文件: {fullFilePath}")
                    End If

                    If File.Exists(fullFilePath) Then
                        ' 根据文件扩展名选择合适的解析方法
                        Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                        Dim fileContentResult As FileContentResult = Nothing

                        Select Case fileExtension
                            Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                fileContentResult = ParseFile(fullFilePath)
                            Case ".docx", ".doc"
                                fileContentResult = ParseFile(fullFilePath)
                            Case ".csv", ".txt"
                                fileContentResult = ParseTextFile(fullFilePath)
                            Case Else
                                fileContentResult = New FileContentResult With {
                            .FileName = Path.GetFileName(fullFilePath),
                            .FileType = "Unknown",
                            .ParsedContent = $"[不支持的文件类型: {fileExtension}]"
                        }
                        End Select

                        If fileContentResult IsNot Nothing Then
                            parsedFiles.Add(fileContentResult)
                            fileContentBuilder.AppendLine($"文件名: {fileContentResult.FileName}")
                            fileContentBuilder.AppendLine($"文件内容:")
                            fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                            fileContentBuilder.AppendLine("---")
                        End If
                    Else
                        fileContentBuilder.AppendLine($"文件 '{Path.GetFileName(filePath)}' 未找到或路径无效")
                        Debug.WriteLine($"文件未找到: {fullFilePath}")
                        ' 尝试列出当前目录中的文件，用于调试
                        If Directory.Exists(currentWorkingDir) Then
                            Dim filesInDir = Directory.GetFiles(currentWorkingDir)
                            Debug.WriteLine($"当前目录中的文件: {String.Join(", ", filesInDir.Select(Function(f) Path.GetFileName(f)))}")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                    fileContentBuilder.AppendLine($"处理文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}")
                    fileContentBuilder.AppendLine("---")
                End Try
            Next

            fileContentBuilder.AppendLine("--- 文件内容结束 ---" & vbCrLf)
        End If

        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If fileContentBuilder.Length > 0 Then
            finalMessageToLLM &= fileContentBuilder.ToString()
        End If

        stopReaderStream = False ' Reset stop flag before sending new message
        SendChatMessage(finalMessageToLLM)
    End Sub

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language, preview)
    End Sub


    ' 抽象方法，由子类实现
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ' 通用的文本文件解析方法
    ' 通用的文本文件解析方法
    Protected Function ParseTextFile(filePath As String) As FileContentResult
        Try
            Dim extension As String = Path.GetExtension(filePath).ToLower()

            ' 对 CSV 文件使用专门的处理逻辑
            If extension = ".csv" Then
                Return ParseCsvFile(filePath)
            End If

            ' 对普通文本文件进行编码检测
            Dim encoding As Encoding = DetectFileEncoding(filePath)
            Dim content As String = File.ReadAllText(filePath, encoding)

            Dim result As New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Text",
            .ParsedContent = content,
            .RawData = content
        }
            Return result
        Catch ex As Exception
            Debug.WriteLine($"Error parsing text file: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Text",
            .ParsedContent = $"[解析文本文件时出错: {ex.Message}]"
        }
        End Try
    End Function

    ' 专门用于解析 CSV 文件的方法
    Protected Function ParseCsvFile(filePath As String) As FileContentResult
        Try
            ' 检测文件编码
            Dim encoding As Encoding = DetectFileEncoding(filePath)

            ' 用检测到的编码读取内容
            Dim csvContent As String = File.ReadAllText(filePath, encoding)

            ' 创建一个格式化的 CSV 内容
            Dim formattedContent As New StringBuilder()
            formattedContent.AppendLine($"CSV 文件: {Path.GetFileName(filePath)} (编码: {encoding.EncodingName})")
            formattedContent.AppendLine()

            ' 分析 CSV 数据结构
            Dim rows As String() = csvContent.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            If rows.Length > 0 Then
                ' 检测分隔符，可能是逗号、分号、制表符等
                Dim delimiter As Char = DetectCsvDelimiter(rows(0))

                ' 获取列数，用于后续限制数据量
                Dim columns As String() = rows(0).Split(delimiter)
                Dim columnCount As Integer = columns.Length

                ' 添加表头（如果存在）
                formattedContent.AppendLine("表头:")
                formattedContent.AppendLine(FormatCsvRow(rows(0), delimiter))
                formattedContent.AppendLine()

                ' 添加数据行（限制行数，避免数据太多）
                Dim maxRows As Integer = Math.Min(rows.Length, 25) ' 最多显示25行
                formattedContent.AppendLine("数据:")

                For i As Integer = 1 To maxRows - 1
                    formattedContent.AppendLine(FormatCsvRow(rows(i), delimiter))
                Next

                ' 如果有更多行，添加提示
                If rows.Length > maxRows Then
                    formattedContent.AppendLine("...")
                    formattedContent.AppendLine($"[文件包含 {rows.Length} 行，仅显示前 {maxRows - 1} 行数据]")
                End If
            Else
                formattedContent.AppendLine("[CSV 文件为空]")
            End If

            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "CSV",
            .ParsedContent = formattedContent.ToString(),
            .RawData = csvContent
        }
        Catch ex As Exception
            Debug.WriteLine($"Error parsing CSV file: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "CSV",
            .ParsedContent = $"[解析 CSV 文件时出错: {ex.Message}]"
        }
        End Try
    End Function

    ' 格式化 CSV 行数据，使其更易读
    Private Function FormatCsvRow(row As String, delimiter As Char) As String
        Dim fields As String() = row.Split(delimiter)
        Dim formattedRow As New StringBuilder()

        For i As Integer = 0 To fields.Length - 1
            Dim field As String = fields(i).Trim(""""c) ' 移除引号
            If i < fields.Length - 1 Then
                formattedRow.Append($"{field} | ")
            Else
                formattedRow.Append(field)
            End If
        Next

        Return formattedRow.ToString()
    End Function

    ' 检测 CSV 文件的分隔符
    Private Function DetectCsvDelimiter(sampleLine As String) As Char
        ' 常见的 CSV 分隔符
        Dim possibleDelimiters As Char() = {","c, ";"c, vbTab, "|"c}
        Dim bestDelimiter As Char = ","c ' 默认使用逗号
        Dim maxCount As Integer = 0

        ' 检查每个可能的分隔符
        For Each delimiter In possibleDelimiters
            Dim count As Integer = sampleLine.Count(Function(c) c = delimiter)
            If count > maxCount Then
                maxCount = count
                bestDelimiter = delimiter
            End If
        Next

        Return bestDelimiter
    End Function

    ' 检测文件编码
    Private Function DetectFileEncoding(filePath As String) As Encoding
        ' 首先，我们尝试从 BOM (Byte Order Mark) 检测编码
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                ' 读取前几个字节来检测 BOM
                Dim bom(3) As Byte
                Dim bytesRead As Integer = fs.Read(bom, 0, bom.Length)

                ' 检查是否有 BOM
                If bytesRead >= 3 AndAlso bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                    ' UTF-8 with BOM
                    Return New UTF8Encoding(True)
                ElseIf bytesRead >= 2 AndAlso bom(0) = &HFE AndAlso bom(1) = &HFF Then
                    ' UTF-16 (Big Endian)
                    Return Encoding.BigEndianUnicode
                ElseIf bytesRead >= 2 AndAlso bom(0) = &HFF AndAlso bom(1) = &HFE Then
                    ' UTF-16 (Little Endian)
                    If bytesRead >= 4 AndAlso bom(2) = 0 AndAlso bom(3) = 0 Then
                        ' UTF-32 (Little Endian)
                        Return Encoding.UTF32
                    Else
                        ' UTF-16 (Little Endian)
                        Return Encoding.Unicode
                    End If
                ElseIf bytesRead >= 4 AndAlso bom(0) = 0 AndAlso bom(1) = 0 AndAlso bom(2) = &HFE AndAlso bom(3) = &HFF Then
                    ' UTF-32 (Big Endian)
                    Return New UTF32Encoding(True, True)
                End If
            End Using

            ' 定义Unicode替换字符，用于检测无效字符
            Dim unicodeReplacementChar As Char = ChrW(&HFFFD) ' Unicode 替换字符 U+FFFD

            ' 针对中文文件，优先尝试 GB18030/GBK 编码
            Dim fileExtension As String = Path.GetExtension(filePath).ToLower()
            If fileExtension = ".csv" Then
                ' 首先尝试 GB18030/GBK 编码，这在中文环境下非常常见
                Try
                    ' 读取部分文件内容
                    Dim csvSampleBytes As Byte() = New Byte(4095) {}  ' 读取前 4KB
                    Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                        fs.Read(csvSampleBytes, 0, csvSampleBytes.Length)
                    End Using

                    ' 尝试用 GB18030 解码
                    Dim gbkEncoding As Encoding = Encoding.GetEncoding("GB18030")
                    Dim gbkText As String = gbkEncoding.GetString(csvSampleBytes)

                    ' 检查解码后的文本是否符合 CSV 格式的特征（包含逗号和换行符）
                    If gbkText.Contains(",") AndAlso (gbkText.Contains(vbCr) OrElse gbkText.Contains(vbLf)) Then
                        ' 如果包含逗号和换行符，可能是有效的 CSV
                        Dim invalidCharCount As Integer = gbkText.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)
                        Dim totalCharCount As Integer = gbkText.Length

                        ' 允许有少量不可识别字符
                        If invalidCharCount <= totalCharCount * 0.05 Then ' 允许5%的不可识别字符
                            Return gbkEncoding
                        End If
                    End If
                Catch ex As Exception
                    ' 忽略错误，继续尝试其他编码
                    Debug.WriteLine($"尝试 GB18030 编码时出错: {ex.Message}")
                End Try
            End If

            ' 尝试几种常见的编码
            Dim encodingsToTry As Encoding() = {
            New UTF8Encoding(False),        ' UTF-8 without BOM
            Encoding.GetEncoding("GB18030"), ' 中文编码，涵盖简体中文
            Encoding.Default                ' 系统默认编码
        }

            ' 读取文件的前几行样本
            Dim generalSampleBytes As Byte() = New Byte(4095) {}  ' 读取前 4KB
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                fs.Read(generalSampleBytes, 0, generalSampleBytes.Length)
            End Using

            Dim bestEncoding As Encoding = encodingsToTry(0) ' 默认使用第一个编码
            Dim leastInvalidCharCount As Integer = Integer.MaxValue

            ' 尝试每种编码，选择产生最少无效字符的编码
            For Each enc In encodingsToTry
                Try
                    Dim sample As String = enc.GetString(generalSampleBytes)
                    ' 计算问号和替换字符的数量作为无效字符的指标
                    Dim invalidCharCount As Integer = sample.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)

                    ' 如果这个编码产生的无效字符更少
                    If invalidCharCount < leastInvalidCharCount Then
                        leastInvalidCharCount = invalidCharCount
                        bestEncoding = enc

                        ' 如果没有无效字符，立即使用这个编码
                        If invalidCharCount = 0 Then
                            Exit For
                        End If
                    End If
                Catch ex As Exception
                    ' 忽略解码错误，尝试下一个编码
                    Continue For
                End Try
            Next

            ' 使用产生最少无效字符的编码
            Return bestEncoding

        Catch ex As Exception
            Debug.WriteLine($"检测文件编码时出错: {ex.Message}")
            ' 出错时使用系统默认编码
            Return Encoding.Default
        End Try
    End Function

    Protected MustOverride Function GetApplication() As ApplicationInfo
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean)
    Protected MustOverride Function RunCode(vbaCode As String)

    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法
    Private Sub ExecuteCode(code As String, language As String, preview As Boolean)
        ' 根据语言类型执行不同的操作
        Dim lowerLang As String = language.ToLower()

        If lowerLang.Contains("vbnet") OrElse lowerLang.Contains("vbscript") OrElse lowerLang.Contains("vba") Then
            ' 执行 VBA 代码 (简化匹配逻辑: 包含vb或vba的都识别为VBA)
            ExecuteVBACode(code, preview)
        ElseIf lowerLang.Contains("js") OrElse lowerLang.Contains("javascript") Then
            ' 执行 JavaScript 代码
            ExecuteJavaScript(code, preview)
        ElseIf lowerLang.Contains("excel") OrElse lowerLang.Contains("formula") OrElse lowerLang.Contains("function") Then
            ' 执行 Excel 函数/公式
            ExecuteExcelFormula(code, preview)

            'Case "sql", "language-sql"
            '    ' 执行 SQL 查询
            '    ExecuteSqlQuery(code, preview)
            'Case "powerquery", "m", "language-powerquery", "language-m"
            '    ' 执行 PowerQuery/M 语言
            '    ExecutePowerQuery(code, preview)
            'Case "python", "py", "language-python"
            '    ' 执行 Python 代码
            '    ExecutePython(code, preview)
        Else
            GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
        End If

    End Sub

    ' 执行JavaScript代码 - 专注于操作Office/WPS对象模型，支持Office JS API风格代码
    Protected Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
        Try
            ' 获取Office应用对象
            Dim appObject As Object = GetOfficeApplicationObject()
            If appObject Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取Office应用程序对象")
                Return False
            End If

            ' 检测是否是Office JS API风格的代码
            Dim isOfficeJsApiStyle As Boolean = jsCode.Contains("getActiveWorksheet") OrElse
                                            jsCode.Contains("getUsedRange") OrElse
                                            jsCode.Contains("getValues") OrElse
                                            jsCode.Contains("setValues")

            ' 创建脚本控制引擎
            Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
            scriptEngine.Language = "JScript"

            ' 判断是WPS还是Microsoft Office
            Dim isWPS As Boolean = False
            Try
                Dim appName As String = appObject.Name
                isWPS = appName.Contains("WPS")
            Catch ex As Exception
                isWPS = False
            End Try

            ' 将Office应用对象暴露给脚本环境
            scriptEngine.AddObject("app", appObject, True)

            ' 添加适配层代码
            Dim adapterCode As String = "
        // Office JS API 适配层
        var Office = {
            isWPS: " & isWPS.ToString().ToLower() & ",
            app: app,
            context: {
                workbook: {
                    // 适配 Office JS API 方法到 COM 对象
                    getActiveWorksheet: function() {
                        return {
                            sheet: app.ActiveSheet,
                            getUsedRange: function() {
                                var usedRange = this.sheet.UsedRange;
                                return {
                                    range: usedRange,
                                    getValues: function() {
                                        var values = [];
                                        var rows = this.range.Rows.Count;
                                        var cols = this.range.Columns.Count;
                                        
                                        for(var i = 1; i <= rows; i++) {
                                            var rowValues = [];
                                            for(var j = 1; j <= cols; j++) {
                                                var cellValue = this.range.Cells(i, j).Value;
                                                rowValues.push(cellValue);
                                            }
                                            values.push(rowValues);
                                        }
                                        return values;
                                    },
                                    setValues: function(values) {
                                        if(!values || values.length === 0) return;
                                        
                                        for(var i = 0; i < values.length; i++) {
                                            var row = values[i];
                                            for(var j = 0; j < row.length; j++) {
                                                try {
                                                    this.range.Cells(i+1, j+1).Value = row[j];
                                                } catch(e) {
                                                    // 忽略单元格设置错误
                                                }
                                            }
                                        }
                                    }
                                };
                            }
                        };
                    }
                }
            },
            // 日志函数
            log: function(message) { 
                return '输出: ' + message; 
            }
        };
        
        // Office JS API 主函数适配器
        function executeOfficeJsApi(codeFunc) {
            var workbook = Office.context.workbook;
            if(typeof codeFunc === 'function') {
                try {
                    return codeFunc(workbook);
                } catch(e) {
                    return 'Office JS API 执行错误: ' + e.message;
                }
            }
            return 'Invalid function';
        }
        "

            ' 预执行适配层代码
            scriptEngine.ExecuteStatement(adapterCode)

            ' 构建执行代码，根据代码类型选择不同的执行方式
            Dim wrappedCode As String

            If isOfficeJsApiStyle Then
                ' 如果是Office JS API风格，使用适配层执行
                wrappedCode = "
            try {
                // 将用户代码包装为函数
                var userFunc = function(workbook) {
                    " & jsCode & "
                };
                
                // 使用适配器执行
                executeOfficeJsApi(userFunc);
                return 'Office JS API 代码执行成功';
            } catch(e) {
                return 'Office JS API 执行错误: ' + e.message;
            }
            "
            Else
                ' 普通JavaScript代码
                wrappedCode = "
            try {
                // 用户代码开始
                " & jsCode & "
                // 用户代码结束
                return '代码执行成功';
            } catch(e) {
                return '执行错误: ' + e.message;
            }
            "
            End If

            ' 执行JavaScript代码并获取结果
            Dim result As String = scriptEngine.Eval(wrappedCode)
            GlobalStatusStrip.ShowInfo(result)

            Return True
        Catch ex As Exception
            MessageBox.Show("执行JavaScript代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function



    ' 添加清除特定 sheetName 的方法
    Public Async Sub ClearSelectedContentBySheetName(sheetName As String)
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
        $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})"
    )
    End Sub


    ' 抽象方法 - 获取Office应用程序对象
    Protected MustOverride Function GetOfficeApplicationObject() As Object

    ' 执行Excel公式或函数 - 基类通用实现
    Protected Function ExecuteExcelFormula(formulaCode As String, preview As Boolean) As Boolean
        Try
            ' 获取应用程序信息
            Dim appInfo As ApplicationInfo = GetApplication()

            ' 去除可能的等号前缀
            If formulaCode.StartsWith("=") Then
                formulaCode = formulaCode.Substring(1)
            End If

            ' 根据应用类型处理
            If appInfo.Type = OfficeApplicationType.Excel Then
                ' 对于Excel，使用Evaluate方法
                Dim result As Boolean = EvaluateFormula(formulaCode, preview)
                GlobalStatusStrip.ShowInfo("公式执行结果: " & result.ToString())
                Return True
            Else
                ' 其他应用不支持直接执行Excel公式
                GlobalStatusStrip.ShowWarning("Excel公式执行仅支持Excel环境")
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("执行Excel公式时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 虚方法 - 评估Excel公式（只有Excel子类会实现）
    Protected Overridable Function EvaluateFormula(formula As String, preview As Boolean) As Boolean
        ' 默认实现返回Nothing
        Return True
    End Function

    ' 执行前端传来的 VBA 代码片段
    Protected Function ExecuteVBACode(vbaCode As String, preview As Boolean)

        If preview Then
            ' 返回是否需要执行，accept-True，reject-False
            If Not RunCodePreview(vbaCode, preview) Then
                Return True
            End If
            ' 如果预览模式，直接返回
        End If

        ' 获取 VBA 项目
        Dim vbProj As VBProject = GetVBProject()

        ' 添加空值检查
        If vbProj Is Nothing Then
            Return False
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
    End Function


    ' 检查代码是否包含过程声明
    Public Function ContainsProcedureDeclaration(code As String) As Boolean
        ' 使用简单的正则表达式检查是否包含 Sub 或 Function 声明
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function


    ' 查找模块中的第一个过程名
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
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            GlobalStatusStrip.ShowWarning("请先配置大模型Api！")
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        If String.IsNullOrWhiteSpace(question) Then
            GlobalStatusStrip.ShowWarning("请输入问题！")

            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        Dim uuid As String = Guid.NewGuid().ToString()

        Try
            Dim requestBody As String = CreateRequestBody(question)
            Await SendHttpRequestStream(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody, StripQuestion(question))
            Await SaveFullWebPageAsync2()
        Catch ex As Exception
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try

    End Function

    Private Sub ManageHistoryMessageSize()
        ' 如果历史消息数超过限制，有一条system，所以+1
        While historyMessageData.Count > contextLimit + 1
            ' 保留系统消息（第一条消息）
            If historyMessageData.Count > 1 Then
                ' 移除第二条消息（最早的非系统消息）
                historyMessageData.RemoveAt(1)
            End If
        End While
    End Sub


    Private Function StripQuestion(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")
        Return result
    End Function

    Private Function CreateRequestBody(question As String) As String
        Dim result As String = StripQuestion(question)

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

        ' 管理历史消息大小
        ManageHistoryMessageSize()

        ' 添加历史消息
        For Each message In historyMessageData
            messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{message.content}""}}")
        Next


        ' 添加MCP工具信息（如果有）
        Dim toolsArray As JArray = Nothing
        Dim chatSettings As New ChatSettings(GetApplication())

        ' 如果有启用的MCP连接
        If ChatSettings.EnabledMcpList IsNot Nothing AndAlso ChatSettings.EnabledMcpList.Count > 0 Then
            toolsArray = New JArray()

            ' 加载所有MCP连接
            Dim connections = MCPConnectionManager.LoadConnections()

            ' 找到启用的连接
            ' 找到启用的连接
            For Each mcpName In chatSettings.EnabledMcpList
                ' 使用IsActive替代Enabled
                Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpName AndAlso c.IsActive)
                If connection IsNot Nothing Then
                    ' 从连接配置中获取已保存的工具列表
                    If connection.Tools IsNot Nothing AndAlso connection.Tools.Count > 0 Then
                        ' 将所有工具添加到工具数组
                        For Each toolObj In connection.Tools
                            toolsArray.Add(toolObj)
                        Next
                        Debug.WriteLine($"从连接 '{connection.Name}' 加载了 {connection.Tools.Count} 个工具")
                    Else
                        ' 如果连接中没有保存工具信息，则使用通用的mcp_call工具
                        Dim toolObj = New JObject()
                        toolObj("type") = "function"
                        toolObj("function") = New JObject()
                        toolObj("function")("name") = "mcp_call"
                        toolObj("function")("description") = $"Call MCP tool through {connection.Name} connection"

                        ' 添加参数架构
                        toolObj("function")("parameters") = New JObject()
                        toolObj("function")("parameters")("type") = "object"
                        toolObj("function")("parameters")("properties") = New JObject()

                        ' 工具名称参数
                        toolObj("function")("parameters")("properties")("tool_name") = New JObject()
                        toolObj("function")("parameters")("properties")("tool_name")("type") = "string"
                        toolObj("function")("parameters")("properties")("tool_name")("description") = "The name of the MCP tool to call"

                        ' 工具参数
                        toolObj("function")("parameters")("properties")("arguments") = New JObject()
                        toolObj("function")("parameters")("properties")("arguments")("type") = "object"
                        toolObj("function")("parameters")("properties")("arguments")("description") = "The arguments to pass to the MCP tool"

                        ' 添加必需参数
                        toolObj("function")("parameters")("required") = New JArray({"tool_name", "arguments"})

                        ' 添加到工具数组
                        toolsArray.Add(toolObj)
                        Debug.WriteLine($"连接 '{connection.Name}' 没有保存工具信息，使用通用mcp_call工具")
                    End If
                End If
            Next
        End If

        ' 构建 JSON 请求体
        Dim messagesJson = String.Join(",", messages)

        ' 如果有工具，添加到请求中
        If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
            Dim toolsJson = toolsArray.ToString(Formatting.None)
            ' 注意这里没有给tools加额外的引号
            Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""tools"": {toolsJson}, ""messages"": [{messagesJson}], ""stream"": true}}"
            Return requestBody
        Else
            ' 直接使用JSON数组符号
            Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"",  ""messages"": [{messagesJson}], ""stream"": true}}"
            Return requestBody
        End If

    End Function


    ' 添加处理MCP工具调用的方法
    Private Async Function HandleMcpToolCall(toolName As String, arguments As JObject, mcpConnectionName As String) As Task(Of JObject)
        Try
            Debug.WriteLine($"开始处理MCP工具调用: 工具={toolName}, 连接={mcpConnectionName}")

            ' 加载MCP连接
            Dim connections = MCPConnectionManager.LoadConnections()
            ' 注意这里使用isActive而不是Enabled
            Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpConnectionName AndAlso c.IsActive)

            If connection Is Nothing Then
                Return CreateErrorResponse($"MCP连接 '{mcpConnectionName}' 未找到或未启用。可用连接: {String.Join(", ", connections.Where(Function(c) c.IsActive).Select(Function(c) c.Name))}")
            End If

            Debug.WriteLine($"找到MCP连接: {connection.Name}, URL: {connection.Url}")

            ' 创建MCP客户端
            Using client As New StreamJsonRpcMCPClient()
                Try
                    ' 配置客户端
                    Await client.ConfigureAsync(connection.Url)
                    Debug.WriteLine("MCP客户端配置完成")

                    ' 初始化连接
                    Dim initResult = Await client.InitializeAsync()
                    If Not initResult.Success Then
                        Return CreateErrorResponse($"初始化MCP连接失败: {initResult.ErrorMessage}。连接URL: {connection.Url}")
                    End If

                    Debug.WriteLine("MCP连接初始化成功")

                    ' 调用工具
                    Debug.WriteLine($"开始调用工具: {toolName}, 参数: {arguments.ToString()}")
                    Dim result = Await client.CallToolAsync(toolName, arguments)

                    ' 处理结果
                    If result.IsError Then
                        Return CreateErrorResponse($"调用MCP工具 '{toolName}' 失败: {result.ErrorMessage}")
                    End If

                    Debug.WriteLine($"工具调用成功，返回内容数量: {result.Content?.Count}")

                    ' 创建成功响应
                    Dim responseObj = New JObject()

                    ' 添加内容数组
                    Dim contentArray = New JArray()
                    If result.Content IsNot Nothing Then
                        For Each content In result.Content
                            Dim contentObj = New JObject()
                            contentObj("type") = content.Type

                            If Not String.IsNullOrEmpty(content.Text) Then
                                contentObj("text") = content.Text
                            End If

                            If Not String.IsNullOrEmpty(content.Data) Then
                                contentObj("data") = content.Data
                            End If

                            If Not String.IsNullOrEmpty(content.MimeType) Then
                                contentObj("mimeType") = content.MimeType
                            End If

                            contentArray.Add(contentObj)
                        Next
                    End If

                    responseObj("content") = contentArray
                    Return responseObj

                Catch clientEx As Exception
                    Debug.WriteLine($"MCP客户端操作失败: {clientEx.Message}")
                    Return CreateErrorResponse($"MCP客户端操作失败: {clientEx.Message}。详细信息: {clientEx.ToString()}")
                End Try
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HandleMcpToolCall整体异常: {ex.Message}")
            Return CreateErrorResponse($"MCP工具调用出现异常: {ex.Message}。工具: {toolName}, 连接: {mcpConnectionName}。堆栈跟踪: {ex.StackTrace}")
        End Try
    End Function

    ' 创建错误响应
    Private Function CreateErrorResponse(errorMessage As String) As JObject
        Dim responseObj = New JObject()
        responseObj("isError") = True
        responseObj("errorMessage") = errorMessage
        responseObj("timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Debug.WriteLine($"创建错误响应: {errorMessage}")
        Return responseObj
    End Function
    ' 添加一个结构来存储token信息
    Private Structure TokenInfo
        Public PromptTokens As Integer
        Public CompletionTokens As Integer
        Public TotalTokens As Integer
    End Structure

    Private totalTokens As Integer = 0
    Private lastTokenInfo As Nullable(Of TokenInfo)

    ' 用于累加当前会话中所有API调用的token消耗（mcp多次消耗的情况）
    Private currentSessionTotalTokens As Integer = 0

    ' 用于跟踪待处理的异步任务
    Private _pendingMcpTasks As Integer = 0
    Private _mainStreamCompleted As Boolean = False
    Private _finalUuid As String = String.Empty
    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String, originQuestion As String) As Task

        ' 组装ai答复头部
        Dim uuid As String = Guid.NewGuid().ToString()

        _finalUuid = uuid
        _mainStreamCompleted = False
        _pendingMcpTasks = 0

        ' 重置当前会话的token累加器
        currentSessionTotalTokens = 0

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
                            Dim buffer(102300) As Char
                            Dim readCount As Integer
                            Do
                                ' 检查是否需要停止读取
                                If stopReaderStream Then
                                    Debug.WriteLine("[Stream] 用户手动停止流读取")
                                    ' 清空当前缓冲区
                                    _currentMarkdownBuffer.Clear()
                                    allMarkdownBuffer.Clear()
                                    ' 停止读取并退出循环
                                    Exit Do
                                End If
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunk As String = New String(buffer, 0, readCount)
                                ' 如果chunk不是以data开头，则跳过
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)
                                ' 判断stringBuilder是否以'}'结尾，如果是则解析
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), uuid, originQuestion)
                                    stringBuilder.Clear()
                                End If
                            Loop
                            Debug.WriteLine("[Stream] 流接收完成")
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 请求过程中出错: {ex.ToString()}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 标记主流已完成
            _mainStreamCompleted = True

            ' 使用累加后的总token数（包含主调用和MCP润色调用）
            Dim finalTokens As Integer = currentSessionTotalTokens
            If lastTokenInfo.HasValue Then
                ' 添加主要调用的tokens
                finalTokens += lastTokenInfo.Value.TotalTokens
                currentSessionTotalTokens += lastTokenInfo.Value.TotalTokens
            End If

            Debug.WriteLine($"finally 主要调用tokens: {If(lastTokenInfo.HasValue, lastTokenInfo.Value.TotalTokens, 0)}")
            Debug.WriteLine($"finally 当前会话总tokens: {currentSessionTotalTokens}")

            ' 检查是否可以完成处理
            CheckAndCompleteProcessing()


            ' 记录全局上下文中，方便后续使用
            Dim answer = New HistoryMessage() With {
                .role = "assistant",
                .content = allMarkdownBuffer.ToString()
            }
            historyMessageData.Add(answer)
            ' 管理历史消息大小
            ManageHistoryMessageSize()

            allMarkdownBuffer.Clear()
            ' 重置token信息
            lastTokenInfo = Nothing
        End Try
    End Function


    ' 检查并完成处理
    Private Sub CheckAndCompleteProcessing()
        Debug.WriteLine($"CheckAndCompleteProcessing: 主流完成={_mainStreamCompleted}, 待处理MCP任务={_pendingMcpTasks}")

        ' 只有在主流完成且没有待处理的MCP任务时才调用完成函数
        If _mainStreamCompleted AndAlso _pendingMcpTasks = 0 Then
            Debug.WriteLine("所有处理完成，调用 processStreamComplete")
            ExecuteJavaScriptAsyncJS($"processStreamComplete('{_finalUuid}',{currentSessionTotalTokens});")
        End If
    End Sub

    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' 启用表格、代码块等扩展
    .Build()                        ' 构建不可变管道

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()

    ' 用于收集工具调用参数的变量
    Private _pendingToolCalls As New Dictionary(Of String, JObject) ' 按ID存储未完成的工具调用
    Private _completedToolCalls As New List(Of JObject) ' 存储已完成的工具调用


    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String, originQuestion As String)
        Try
            Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                line = line.Trim()
                If line = "[DONE]" Then
                    ' 在流结束时处理所有完成的工具调用
                    If _pendingToolCalls.Count > 0 Then
                        Debug.WriteLine("[DONE] 时发现未处理的工具调用，开始处理")
                        ProcessCompletedToolCalls(uuid, originQuestion)
                    End If
                    FlushBuffer("content", uuid) ' 最后刷新缓冲区
                    Return
                End If
                If line = "" Then
                    Continue For
                End If

                Debug.Print(line)
                Dim jsonObj As JObject = JObject.Parse(line)

                ' 获取token信息 - 只保存最后一个响应块的usage信息
                Dim usage = jsonObj("usage")
                If usage IsNot Nothing AndAlso usage.Type = JTokenType.Object Then
                    lastTokenInfo = New TokenInfo With {
                    .PromptTokens = CInt(usage("prompt_tokens")),
                    .CompletionTokens = CInt(usage("completion_tokens")),
                    .TotalTokens = CInt(usage("total_tokens"))
                }
                End If

                Dim reasoning_content As String = jsonObj("choices")(0)("delta")("reasoning_content")?.ToString()
                If Not String.IsNullOrEmpty(reasoning_content) Then
                    _currentMarkdownBuffer.Append(reasoning_content)
                    FlushBuffer("reasoning", uuid)
                End If

                Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()
                If Not String.IsNullOrEmpty(content) Then
                    _currentMarkdownBuffer.Append(content)
                    FlushBuffer("content", uuid)
                End If

                ' 检查是否有工具调用
                Dim choices = jsonObj("choices")
                If choices IsNot Nothing AndAlso choices.Count > 0 Then
                    Dim choice = choices(0)
                    Dim delta = choice("delta")
                    Dim finishReason = choice("finish_reason")?.ToString()

                    ' 收集工具调用数据
                    If delta IsNot Nothing Then
                        Dim toolCalls = delta("tool_calls")
                        If toolCalls IsNot Nothing AndAlso toolCalls.Count > 0 Then
                            CollectToolCallData(toolCalls, originQuestion)
                        End If
                    End If

                    ' 当finish_reason为tool_calls时，说明所有工具调用数据已接收完毕
                    If finishReason = "tool_calls" Then
                        Debug.WriteLine("检测到 finish_reason = tool_calls，开始处理工具调用")
                        ProcessCompletedToolCalls(uuid, originQuestion)
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 数据处理失败: {ex.Message}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 收集工具调用数据
    Private Sub CollectToolCallData(toolCalls As JArray, originQuestion As String)
        Try
            For Each toolCall In toolCalls
                Dim toolIndex = toolCall("index")?.Value(Of Integer)()
                Dim toolId = toolCall("id")?.ToString()

                ' 统一使用index作为主键，因为index是唯一且连续的
                Dim toolKey As String = $"tool_{toolIndex}"

                ' 如果是新的工具调用，创建新的条目
                If Not _pendingToolCalls.ContainsKey(toolKey) Then
                    _pendingToolCalls(toolKey) = New JObject()
                    ' 保存真实的ID，但使用index作为内部键
                    _pendingToolCalls(toolKey)("realId") = If(String.IsNullOrEmpty(toolId), toolKey, toolId)
                    _pendingToolCalls(toolKey)("index") = toolIndex
                    _pendingToolCalls(toolKey)("type") = toolCall("type")?.ToString()
                    _pendingToolCalls(toolKey)("function") = New JObject()
                    _pendingToolCalls(toolKey)("function")("name") = ""
                    _pendingToolCalls(toolKey)("function")("arguments") = ""
                    _pendingToolCalls(toolKey)("processed") = False
                End If

                Dim currentTool = _pendingToolCalls(toolKey)

                ' 累积函数名称
                Dim functionName = toolCall("function")("name")?.ToString()
                If Not String.IsNullOrEmpty(functionName) Then
                    currentTool("function")("name") = functionName
                    Debug.WriteLine($"设置工具名称: Key={toolKey}, Name={functionName}")
                End If

                ' 累积参数
                Dim arguments = toolCall("function")("arguments")?.ToString()
                If Not String.IsNullOrEmpty(arguments) Then
                    Dim currentArgs = currentTool("function")("arguments").ToString()
                    currentTool("function")("arguments") = currentArgs & arguments
                    Debug.WriteLine($"收集工具调用数据: Key={toolKey}, 本次参数片段='{arguments}', 累积后参数长度={currentTool("function")("arguments").ToString().Length}")
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"收集工具调用数据时出错: {ex.Message}")
        End Try
    End Sub

    ' 处理所有已完成的工具调用
    Private Async Sub ProcessCompletedToolCalls(uuid As String, originQuestion As String)
        Try
            If _pendingToolCalls.Count = 0 Then Return

            Debug.WriteLine($"开始处理 {_pendingToolCalls.Count} 个工具调用")

            For Each kvp In _pendingToolCalls
                Dim toolCall = kvp.Value
                Dim toolKey = kvp.Key

                ' 检查是否已经处理过
                If CBool(toolCall("processed")) Then
                    Debug.WriteLine($"工具调用 {toolKey} 已处理，跳过")
                    Continue For
                End If

                Dim toolName = toolCall("function")("name").ToString()
                Dim argumentsStr = toolCall("function")("arguments").ToString()

                ' 验证工具调用是否完整 - 必须同时有名称和参数
                If String.IsNullOrEmpty(toolName) Then
                    Debug.WriteLine($"工具调用 {toolKey} 缺少名称，跳过处理")
                    Continue For
                End If

                ' 如果参数为空，也跳过（除非某些工具真的不需要参数）
                If String.IsNullOrEmpty(argumentsStr) Then
                    Debug.WriteLine($"工具调用 {toolKey} 参数为空，使用空对象")
                End If

                Debug.WriteLine($"处理工具调用: Key={toolKey}, Name={toolName}, Arguments={argumentsStr}")

                ' 标记为已处理，防止重复执行
                toolCall("processed") = True

                ' 验证参数是否为有效JSON
                Dim argumentsObj As JObject = Nothing
                Try
                    If Not String.IsNullOrEmpty(argumentsStr) Then
                        argumentsObj = JObject.Parse(argumentsStr)
                        Debug.WriteLine($"成功解析参数JSON: {argumentsObj.ToString()}")
                    Else
                        argumentsObj = New JObject()
                        Debug.WriteLine("参数为空，使用空对象")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"工具 {toolName} 的参数格式错误: {ex.Message}, 原始参数: {argumentsStr}")

                    ' 通过FlushBuffer向前端显示详细错误
                    Dim errorMessage = $"<br/>**工具调用参数解析错误：**<br/>" &
                                     $"工具名称: {toolName}<br/>" &
                                     $"错误详情: {ex.Message}<br/>" &
                                     $"原始参数: `{argumentsStr}`<br/>"
                    _currentMarkdownBuffer.Append(errorMessage)
                    FlushBuffer("content", uuid)

                    Continue For ' 跳过这个有问题的工具调用
                End Try

                ' 添加消息到界面，说明正在调用工具
                Dim toolCallMessage = $"<br/>**正在调用工具: {toolName}**<br/>参数: `{argumentsObj.ToString(Formatting.None)}`<br/>"
                _currentMarkdownBuffer.Append(toolCallMessage)
                FlushBuffer("content", uuid)

                ' 从设置中获取启用的MCP连接
                Dim chatSettings As New ChatSettings(GetApplication())
                Dim enabledMcpList = chatSettings.EnabledMcpList

                If enabledMcpList IsNot Nothing AndAlso enabledMcpList.Count > 0 Then
                    ' 使用第一个启用的MCP连接
                    Dim mcpConnectionName = enabledMcpList(0)

                    ' 调用工具
                    Dim result = Await HandleMcpToolCall(toolName, argumentsObj, mcpConnectionName)

                    ' 处理结果
                    If result("isError") IsNot Nothing AndAlso CBool(result("isError")) Then
                        ' 通过FlushBuffer显示详细错误信息
                        Dim detailedError = result("content")?.ToString()
                        Dim errorMessage = $"<br/>**工具调用失败：**<br/>" &
                                         $"**工具名称:** {toolName}<br/>" &
                                         $"**连接名称:** {mcpConnectionName}<br/>" &
                                         $"**错误详情:** {detailedError}<br/>" &
                                         $"**调用参数:**<br/>```json{vbCrLf}{argumentsObj.ToString(Formatting.Indented)}{vbCrLf}```<br/>"

                        _currentMarkdownBuffer.Append(errorMessage)
                        FlushBuffer("content", uuid)
                    Else
                        ' 增加待处理任务计数
                        _pendingMcpTasks += 1
                        Debug.WriteLine($"增加MCP任务，当前待处理任务数: {_pendingMcpTasks}")

                        ' 不直接显示结果，而是发送给大模型进行润色
                        Await SendToolResultForFormatting(toolName, argumentsObj, result, uuid, originQuestion)
                    End If
                Else
                    ' 没有启用的MCP连接
                    Dim errorMessage = "<br/>**配置错误：**<br/>没有启用的MCP连接，无法调用工具。请在设置中启用MCP连接。<br/>"
                    _currentMarkdownBuffer.Append(errorMessage)
                    FlushBuffer("content", uuid)
                End If
            Next

            ' 清空已处理的工具调用
            _pendingToolCalls.Clear()
            _completedToolCalls.Clear()

        Catch ex As Exception
            Debug.WriteLine($"处理完成的工具调用时出错: {ex.Message}")

            ' 向前端显示处理错误
            Dim errorMessage = $"<br/>**工具调用处理异常：**<br/>" &
                             $"**错误详情:** {ex.Message}<br/>" &
                             $"**堆栈跟踪:**<br/>```{vbCrLf}{ex.StackTrace}{vbCrLf}```<br/>"
            _currentMarkdownBuffer.Append(errorMessage)
            FlushBuffer("content", uuid)
        End Try
    End Sub

    ' 新增方法：发送工具结果给大模型进行润色
    Private Async Function SendToolResultForFormatting(toolName As String, arguments As JObject, result As JObject, uuid As String, originQuestion As String) As Task
        Try
            ' 准备发送给大模型的消息内容
            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"用户的原始问题：'{originQuestion}' ,但用户使用了 MCP 工具 '{toolName}'，参数为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(arguments.ToString(Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("工具执行结果为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(result.ToString(Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("请将上述结果结合用户的原始问题整理成易于理解的格式，并使用合适的Markdown格式化呈现，突出重要信息。不需要解释工具调用过程，只需要呈现结果。不要重复用户的请求内容。")

            ' 构建请求体
            Dim messagesArray = New JArray()
            Dim systemMessage = New JObject()
            systemMessage("role") = "system"
            systemMessage("content") = "你是一个帮助解释API调用结果的助手。你的任务是将MCP工具返回的JSON结果转换为人类易读的格式，可适当根据用户原始问题作出取舍，并用Markdown呈现，且没有任何一句废话。"

            Dim userMessage = New JObject()
            userMessage("role") = "user"
            userMessage("content") = promptBuilder.ToString()

            messagesArray.Add(systemMessage)
            messagesArray.Add(userMessage)

            Dim requestObj = New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = True

            Dim requestBody = requestObj.ToString(Formatting.None)

            ' 用于存储当前MCP润色调用的token信息
            Dim mcpTokenInfo As Nullable(Of TokenInfo) = Nothing

            ' 发送请求
            Using client As New HttpClient()
                client.Timeout = Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, ConfigSettings.ApiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", ConfigSettings.ApiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    ' 处理流响应
                    Dim formattedBuilder As New StringBuilder()
                    formattedBuilder.AppendLine("<br/>**工具调用结果：**<br/>")

                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim stringBuilder As New StringBuilder()
                            Dim buffer(1023) As Char
                            Dim readCount As Integer

                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do

                                Dim chunk As String = New String(buffer, 0, readCount)
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)

                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    Dim lines As String() = stringBuilder.ToString().Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                                    For Each line In lines
                                        line = line.Trim()
                                        If line = "[DONE]" Then
                                            Continue For
                                        End If
                                        If line = "" Then
                                            Continue For
                                        End If

                                        Try
                                            Dim jsonObj As JObject = JObject.Parse(line)
                                            ' 收集token信息
                                            Dim usage = jsonObj("usage")
                                            If usage IsNot Nothing Then
                                                mcpTokenInfo = New TokenInfo With {
                                                    .PromptTokens = CInt(usage("prompt_tokens")),
                                                    .CompletionTokens = CInt(usage("completion_tokens")),
                                                    .TotalTokens = CInt(usage("total_tokens"))
                                                }
                                                'Debug.WriteLine($"MCP润色调用tokens: {mcpTokenInfo.Value.TotalTokens}")
                                            End If

                                            Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()

                                            If Not String.IsNullOrEmpty(content) Then
                                                formattedBuilder.Append(content)
                                            End If
                                        Catch ex As Exception
                                            ' 忽略解析错误
                                            Debug.WriteLine($"解析工具结果润色响应出错: {ex.Message}")
                                        End Try
                                    Next

                                    stringBuilder.Clear()
                                End If
                            Loop
                        End Using
                    End Using

                    ' 显示格式化后的结果
                    _currentMarkdownBuffer.Append(formattedBuilder.ToString())
                    FlushBuffer("content", uuid)

                    ' 累加MCP润色调用的token消耗
                    If mcpTokenInfo.HasValue Then
                        currentSessionTotalTokens += mcpTokenInfo.Value.TotalTokens
                        Debug.WriteLine($"累加MCP润色tokens: {mcpTokenInfo.Value.TotalTokens}, 当前总tokens: {currentSessionTotalTokens}")
                    End If
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"格式化工具结果时出错: {ex.Message}")

            ' 如果格式化失败，直接显示原始结果
            _currentMarkdownBuffer.Append("\n\n**工具调用结果：**\n\n```json\n")
            _currentMarkdownBuffer.Append(result.ToString(Formatting.Indented))
            _currentMarkdownBuffer.Append("\n```\n")
            FlushBuffer("content", uuid)
        Finally
            ' 减少待处理任务计数
            _pendingMcpTasks -= 1
            Debug.WriteLine($"MCP任务完成，当前待处理任务数: {_pendingMcpTasks}")

            ' 检查是否可以完成处理
            CheckAndCompleteProcessing()
        End Try
    End Function

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

                                        ' 2. 使用正则表达式移除底部输入栏
                                        Dim bottomBarPattern As String = "<div[^>]*id=[""']chat-bottom-bar[""'][^>]*>.*?</div>\s*</div>\s*</div>"
                                        decodedHtml = Regex.Replace(decodedHtml, bottomBarPattern, "", RegexOptions.Singleline)

                                        ' 移除 <script> 标签及其内容
                                        Dim scriptPattern As String = "<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>"
                                        decodedHtml = Regex.Replace(decodedHtml, scriptPattern, String.Empty, RegexOptions.IgnoreCase)
                                        decodedHtml = decodedHtml.Replace("https://officeai.local/css/", "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.7.0/styles/")


                                        ' 重新注入必要的JavaScript代码
                                        Dim essentialScript As String = GetEssentialJavaScript()

                                        ' 在 </body> 标签前插入必要的脚本
                                        If decodedHtml.Contains("</body>") Then
                                            decodedHtml = decodedHtml.Replace("</body>", essentialScript & Environment.NewLine & "</body>")
                                        Else
                                            ' 如果没有 </body> 标签，在末尾添加
                                            decodedHtml &= essentialScript
                                        End If

                                        tcs.SetResult(decodedHtml)
                                    Catch ex As Exception
                                        tcs.SetException(ex)
                                    End Try
                                End Sub)

        Return Await tcs.Task
    End Function

    Private Function GetEssentialJavaScript() As String
        Return "
<script>
// 代码复制功能
function copyCode(button) {
    const codeBlock = button.closest('.code-block');
    const codeElement = codeBlock.querySelector('code');
    const code = codeElement.textContent;

    const textarea = document.createElement('textarea');
    textarea.value = code;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);

    try {
        textarea.select();
        textarea.setSelectionRange(0, 99999);
        document.execCommand('copy');

        const originalText = button.innerHTML;
        button.innerHTML = '已复制';
        setTimeout(() => {
            button.innerHTML = originalText;
        }, 2000);
    } catch (err) {
        console.error('复制失败:', err);
        alert('复制失败');
    } finally {
        document.body.removeChild(textarea);
    }
}

// 聊天消息引用展开/折叠功能
function toggleChatMessageReference(headerElement) {
    const container = headerElement.closest('.chat-message-references');
    if (container) {
        container.classList.toggle('collapsed');
        
        // 更新箭头方向
        const arrow = headerElement.querySelector('.chat-message-reference-arrow');
        if (arrow) {
            arrow.innerHTML = container.classList.contains('collapsed') ? '&#9658;' : '&#9660;';
        }
    }
}

// 页面初始化
document.addEventListener('DOMContentLoaded', function() {
    // 添加代码块点击事件
    document.querySelectorAll('.code-toggle-label').forEach(label => {
        label.onclick = function(e) {
            e.stopPropagation();
            const preElement = this.nextElementSibling;
            if (preElement && preElement.tagName.toLowerCase() === 'pre') {
                preElement.classList.toggle('collapsed');
                this.textContent = preElement.classList.contains('collapsed') ? '点击展开代码' : '点击折叠代码';
            }
        };
    });
    
    // 添加pre元素点击事件
    document.querySelectorAll('pre.collapsible').forEach(preElement => {
        preElement.onclick = function(e) {
            // 忽略代码按钮点击
            if (e.target.closest('.code-button') || e.target.closest('.code-buttons')) {
                return;
            }
            e.stopPropagation();
            this.classList.toggle('collapsed');
            
            const toggleLabel = this.previousElementSibling;
            if (toggleLabel && toggleLabel.classList.contains('code-toggle-label')) {
                toggleLabel.textContent = this.classList.contains('collapsed') ? '点击展开代码' : '点击折叠代码';
            }
        };
    });
    
    // 处理聊天消息引用点击
    document.querySelectorAll('.chat-message-reference-header').forEach(header => {
        header.onclick = function(e) {
            e.preventDefault();
            e.stopPropagation();
            toggleChatMessageReference(this);
        };
    });
    
    // 处理推理容器点击
    document.querySelectorAll('.reasoning-header').forEach(header => {
        header.onclick = function() {
            const container = this.closest('.reasoning-container');
            if (container) {
                container.classList.toggle('collapsed');
            }
        };
    });
});

// 如果DOM已加载完成，立即执行初始化
if (document.readyState !== 'loading') {
    const event = new Event('DOMContentLoaded');
    document.dispatchEvent(event);
}
</script>"
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

        ' 注入辅助脚本
        Dim script As String = "
        window.vsto = {
            executeCode: function(code, language,preview) {
                window.chrome.webview.postMessage({
                    type: 'executeCode',
                    code: code,
                    language: language,
                    executecodePreview: preview
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
            sendMessage: function(payload) {
                let messageToSend;
                if (typeof payload === 'string') {
                    messageToSend = { type: 'sendMessage', value: payload };
                } else {
                    messageToSend = payload;
                }
                window.chrome.webview.postMessage(messageToSend);
                return true;
            },
            saveSettings: function(settingsObject){
                return window.chrome.webview.postMessage({
                    type: 'saveSettings',
                    topicRandomness: settingsObject.topicRandomness,
                    contextLimit: settingsObject.contextLimit,
                    selectedCell: settingsObject.selectedCell,
                    executeCodePreview: settingsObject.executeCodePreview,
                });
            }
        };
    "
        ChatBrowser.ExecuteScriptAsync(script)
    End Sub

    ' 选中内容发送到聊天区
    Public Async Sub AddSelectedContentItem(sheetName As String, address As String)
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