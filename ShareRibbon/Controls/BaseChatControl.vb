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

    ' 存储聊天HTML的文件路径
    Protected ReadOnly ChatHtmlFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
    )

    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If e.IsSuccess Then
            Try
                If ChatBrowser.InvokeRequired Then
                    ' 使用同步的 Invoke 而不是异步的
                    ChatBrowser.Invoke(Sub()
                                           Task.Delay(100).Wait() ' 同步等待
                                           InitializeSettings()

                                           ' 直接在UI线程移除事件处理器
                                           If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                                               RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                                           End If
                                       End Sub)
                Else
                    Task.Delay(100).Wait() ' 同步等待
                    InitializeSettings()

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
            
            var selectElement = document.getElementById('chatMode');
            if (selectElement) {{
                selectElement.value = '{chatSettings.chatMode}';
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
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case "saveSettings"
                    Debug.Print("保存设置")
                    HandleSaveSettings(jsonDoc)
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

    Protected Overridable Sub HandleSaveSettings(jsonDoc As JObject)
        topicRandomness = jsonDoc("topicRandomness")
        contextLimit = jsonDoc("contextLimit")
        selectedCellChecked = jsonDoc("selectedCell")
        settingsScrollChecked = jsonDoc("settingsScroll")
        Dim chatMode As String = jsonDoc("chatMode")
        Dim chatSettings As New ChatSettings(GetApplication())
        ' 保存设置到配置文件
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, chatMode)
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

    ' 在 HandleSendMessage 方法中添加文件内容解析逻辑
    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String
        Dim filePaths As List(Of String) = New List(Of String)()
        Dim selectedContents As List(Of SendMessageReferenceContentItem) = New List(Of SendMessageReferenceContentItem)()

        If messageValue.Type = JTokenType.String Then
            ' Legacy format or simple text message
            question = messageValue.ToString()
            If question = "InnerStopBtn_#" Then
                stopReaderStream = True
                Return
            End If
        ElseIf messageValue.Type = JTokenType.Object Then
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

        ' --- 处理选中的内容 ---
        question = AppendCurrentSelectedContent("--- 用户的问题：" & question & " 。用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---")

        ' --- 文件内容解析逻辑 ---
        Dim fileContentBuilder As New StringBuilder()
        Dim parsedFiles As New List(Of FileContentResult)()

        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
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
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language)
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
    Protected MustOverride Function RunCode(vbaCode As String)
    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法
    Private Sub ExecuteCode(code As String, language As String)
        ' 根据语言类型执行不同的操作
        Select Case language.ToLower()
            Case "vba", "vb", "vbscript", "language-vba", "language-vbscript", "language-vba hljs language-vbscript", "vba hljs language-vbscript"
                ' 执行 VBA 代码
                'ExecuteVBACode(code)
                RunCode(code)
            Case Else
                'MessageBox.Show("不支持的语言类型: " & language)
                GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
        End Select
    End Sub




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

        ' 管理历史消息大小
        ManageHistoryMessageSize()

        ' 添加历史消息
        For Each message In historyMessageData
            messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{message.content}""}}")
        Next

        ' 构建 JSON 请求体
        Dim messagesJson = String.Join(",", messages)
        Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""messages"": [{messagesJson}], ""stream"": true}}"

        Return requestBody
    End Function


    ' 添加一个结构来存储token信息
    Private Structure TokenInfo
        Public PromptTokens As Integer
        Public CompletionTokens As Integer
        Public TotalTokens As Integer
    End Structure

    Private totalTokens As Integer = 0
    Private lastTokenInfo As Nullable(Of TokenInfo)
    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String) As Task

        ' 组装ai答复头部
        Dim uuid As String = Guid.NewGuid().ToString()
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
                                'Debug.WriteLine($"[Stream] 接收到块:{stringBuilder}")
                                ' 判断stringBuilder是否以'}'结尾，如果是则解析
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), uuid)
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
            ' 使用最后一个响应块中的total_tokens
            Dim finalTokens As Integer = If(lastTokenInfo.HasValue, lastTokenInfo.Value.TotalTokens, 0)
            Debug.WriteLine($"finally {finalTokens}")
            ExecuteJavaScriptAsyncJS($"processStreamComplete('{uuid}',{finalTokens});")

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



    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' 启用表格、代码块等扩展
    .Build()                        ' 构建不可变管道

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()



    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String)
        Try
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

                ' 获取token信息 - 只保存最后一个响应块的usage信息
                Dim usage = jsonObj("usage")
                If usage IsNot Nothing Then
                    lastTokenInfo = New TokenInfo With {
                    .PromptTokens = CInt(usage("prompt_tokens")),
                    .CompletionTokens = CInt(usage("completion_tokens")),
                    .TotalTokens = CInt(usage("total_tokens"))
                }
                End If

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

        ' 检查本地HTML文件是否存在 加载本地HTML文件
        'Dim htmlFilePath As String = ChatHtmlFilePath
        'If File.Exists(htmlFilePath) Then
        '    LoadLocalHtmlFile()
        'End If

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