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


    ' ai����ʷ�ظ�
    Protected historyMessageData As New List(Of HistoryMessage)

    Protected loadingPictureBox As PictureBox

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PASTE As Integer = &H302
        If m.Msg = WM_PASTE Then
            ' �ڴ˴���ճ�����������磺
            If Clipboard.ContainsText() Then
                Dim txt As String = Clipboard.GetText()

                'QuestionTextBox.Text &= txt ' ��ճ������ֱ��д�뵱ǰ���λ��
            End If
            ' ������Ϣ���ݸ����࣬�Ӷ����غ�������  
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Protected Async Function InitializeWebView2() As Task
        Try
            ' �Զ����û�����Ŀ¼
            Dim userDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyAppWebView2Cache")

            ' ȷ��Ŀ¼����
            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            ' �ͷ���Դ�ļ�������
            Dim wwwRoot As String = ResourceExtractor.ExtractResources()

            ' ���� WebView2 �Ĵ�������
            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                .UserDataFolder = userDataFolder
            }

            ' ��ʼ�� WebView2
            Await ChatBrowser.EnsureCoreWebView2Async(Nothing)

            ' ȷ�� CoreWebView2 �ѳ�ʼ��
            If ChatBrowser.CoreWebView2 IsNot Nothing Then

                ' ���� WebView2 �İ�ȫѡ��
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True ' ����ʱ���ÿ����߹���

                ' ��������������ӳ�䵽����Ŀ¼
                ChatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "officeai.local", ' ����������
                    wwwRoot,          ' �����ļ���·��
                    CoreWebView2HostResourceAccessKind.Allow  ' �������
                )

                ' �滻ģ���е� {wwwroot} ռλ��
                Dim htmlContent As String = My.Resources.chat_template

                ' ���� HTML ģ��
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)

                ' ���� Marked �ʹ������
                ConfigureMarked()
                ' ��ӵ�������¼�����ȷ����ҳ�������ɺ��ʼ������
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted

            Else
                MessageBox.Show("WebView2 ��ʼ��ʧ�ܣ�CoreWebView2 �����á�", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"��ʼ��ʧ��: {ex.Message}{Environment.NewLine}����: {ex.GetType().Name}{Environment.NewLine}��ջ:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Async Sub InjectScript(scriptContent As String)
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim escapedScript = JsonConvert.SerializeObject(scriptContent)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedScript})")
        Else
            MessageBox.Show("CoreWebView2 δ��ʼ�����޷�ע��ű���", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            MessageBox.Show("CoreWebView2 δ��ʼ�����޷����� Marked��", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Function

    ' �洢����HTML���ļ�·��
    Protected ReadOnly ChatHtmlFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
    )

    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If e.IsSuccess Then
            Try
                If ChatBrowser.InvokeRequired Then
                    ' ʹ��ͬ���� Invoke �������첽��
                    ChatBrowser.Invoke(Sub()
                                           Task.Delay(100).Wait() ' ͬ���ȴ�
                                           InitializeSettings()

                                           ' ֱ����UI�߳��Ƴ��¼�������
                                           If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                                               RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                                           End If
                                       End Sub)
                Else
                    Task.Delay(100).Wait() ' ͬ���ȴ�
                    InitializeSettings()

                    ' ֱ����UI�߳��Ƴ��¼�������
                    If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                        RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"��������¼������г���: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        End If
    End Sub

    Protected Sub InitializeSettings()
        Try
            ' ��������
            Dim chatSettings As New ChatSettings(GetApplication())
            selectedCellChecked = ChatSettings.selectedCellChecked
            contextLimit = ChatSettings.contextLimit
            topicRandomness = ChatSettings.topicRandomness
            settingsScrollChecked = ChatSettings.settingsScrollChecked

            ' �����÷��͵�ǰ��
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
            
            // ͬ�����������checkbox
            document.getElementById('scrollChecked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('selectedCell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
        "
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"��ʼ������ʧ��: {ex.Message}")
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
                    Debug.Print("��������")
                    HandleSaveSettings(jsonDoc)
                Case Else
                    Debug.WriteLine($"δ֪��Ϣ����: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"������Ϣ����: {ex.Message}")
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
        ' �������õ������ļ�
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, chatMode)
    End Sub

    Public Class SendMessageReferenceContentItem
        Public Property id As String
        Public Property sheetName As String
        Public Property address As String
    End Class

    ' �����ļ����ݽ����������
    Public Class FileContentResult
        Public Property FileName As String
        Public Property FileType As String  ' "Excel", "Word", "Text", ��
        Public Property ParsedContent As String  ' ��ʽ���������ַ���
        Public Property RawData As Object  ' ԭʼ���ݣ������ڽ�һ������
    End Class

    ' �� HandleSendMessage ����������ļ����ݽ����߼�
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

            ' ���� selectedContent
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

        ' --- ����ѡ�е����� ---
        question = AppendCurrentSelectedContent("--- �û������⣺" & question & " ���û����ʽ������������õ��ļ�����ͬһĿ¼�����Կ��Է��Ķ�ȡ�� ---")

        ' --- �ļ����ݽ����߼� ---
        Dim fileContentBuilder As New StringBuilder()
        Dim parsedFiles As New List(Of FileContentResult)()

        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            fileContentBuilder.AppendLine(vbCrLf & "--- �������û����õ������ļ����� ---")

            ' ��ȡ��ǰ����Ŀ¼
            Dim currentWorkingDir As String = GetCurrentWorkingDirectory()
            If String.IsNullOrEmpty(currentWorkingDir) Then
                GlobalStatusStrip.ShowWarning("�뱣�浱ǰ�ļ������Ұ������ļ��͵�ǰ�ļ�����ͬһĿ¼�º�����: ")
                Return
            End If

            For Each filePath As String In filePaths
                Try
                    ' ����ļ��Ƿ�Ϊ����·��
                    Dim fullFilePath As String = filePath

                    ' ������Ǿ���·���������ڵ�ǰ����Ŀ¼�²���
                    If Not Path.IsPathRooted(filePath) AndAlso Not String.IsNullOrEmpty(currentWorkingDir) Then
                        fullFilePath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                        Debug.WriteLine($"�����ڹ���Ŀ¼�����ļ�: {fullFilePath}")
                    End If

                    If File.Exists(fullFilePath) Then
                        ' �����ļ���չ��ѡ����ʵĽ�������
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
                            .ParsedContent = $"[��֧�ֵ��ļ�����: {fileExtension}]"
                        }
                        End Select

                        If fileContentResult IsNot Nothing Then
                            parsedFiles.Add(fileContentResult)
                            fileContentBuilder.AppendLine($"�ļ���: {fileContentResult.FileName}")
                            fileContentBuilder.AppendLine($"�ļ�����:")
                            fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                            fileContentBuilder.AppendLine("---")
                        End If
                    Else
                        fileContentBuilder.AppendLine($"�ļ� '{Path.GetFileName(filePath)}' δ�ҵ���·����Ч")
                        Debug.WriteLine($"�ļ�δ�ҵ�: {fullFilePath}")
                        ' �����г���ǰĿ¼�е��ļ������ڵ���
                        If Directory.Exists(currentWorkingDir) Then
                            Dim filesInDir = Directory.GetFiles(currentWorkingDir)
                            Debug.WriteLine($"��ǰĿ¼�е��ļ�: {String.Join(", ", filesInDir.Select(Function(f) Path.GetFileName(f)))}")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                    fileContentBuilder.AppendLine($"�����ļ� '{Path.GetFileName(filePath)}' ʱ����: {ex.Message}")
                    fileContentBuilder.AppendLine("---")
                End Try
            Next

            fileContentBuilder.AppendLine("--- �ļ����ݽ��� ---" & vbCrLf)
        End If

        ' �������շ��͸� LLM ����Ϣ
        Dim finalMessageToLLM As String = question

        ' Ȼ������ļ����ݣ�����У�
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


    ' ���󷽷���������ʵ��
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ' ͨ�õ��ı��ļ���������
    ' ͨ�õ��ı��ļ���������
    Protected Function ParseTextFile(filePath As String) As FileContentResult
        Try
            Dim extension As String = Path.GetExtension(filePath).ToLower()

            ' �� CSV �ļ�ʹ��ר�ŵĴ����߼�
            If extension = ".csv" Then
                Return ParseCsvFile(filePath)
            End If

            ' ����ͨ�ı��ļ����б�����
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
            .ParsedContent = $"[�����ı��ļ�ʱ����: {ex.Message}]"
        }
        End Try
    End Function

    ' ר�����ڽ��� CSV �ļ��ķ���
    Protected Function ParseCsvFile(filePath As String) As FileContentResult
        Try
            ' ����ļ�����
            Dim encoding As Encoding = DetectFileEncoding(filePath)

            ' �ü�⵽�ı����ȡ����
            Dim csvContent As String = File.ReadAllText(filePath, encoding)

            ' ����һ����ʽ���� CSV ����
            Dim formattedContent As New StringBuilder()
            formattedContent.AppendLine($"CSV �ļ�: {Path.GetFileName(filePath)} (����: {encoding.EncodingName})")
            formattedContent.AppendLine()

            ' ���� CSV ���ݽṹ
            Dim rows As String() = csvContent.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            If rows.Length > 0 Then
                ' ���ָ����������Ƕ��š��ֺš��Ʊ����
                Dim delimiter As Char = DetectCsvDelimiter(rows(0))

                ' ��ȡ���������ں�������������
                Dim columns As String() = rows(0).Split(delimiter)
                Dim columnCount As Integer = columns.Length

                ' ��ӱ�ͷ��������ڣ�
                formattedContent.AppendLine("��ͷ:")
                formattedContent.AppendLine(FormatCsvRow(rows(0), delimiter))
                formattedContent.AppendLine()

                ' ��������У�������������������̫�ࣩ
                Dim maxRows As Integer = Math.Min(rows.Length, 25) ' �����ʾ25��
                formattedContent.AppendLine("����:")

                For i As Integer = 1 To maxRows - 1
                    formattedContent.AppendLine(FormatCsvRow(rows(i), delimiter))
                Next

                ' ����и����У������ʾ
                If rows.Length > maxRows Then
                    formattedContent.AppendLine("...")
                    formattedContent.AppendLine($"[�ļ����� {rows.Length} �У�����ʾǰ {maxRows - 1} ������]")
                End If
            Else
                formattedContent.AppendLine("[CSV �ļ�Ϊ��]")
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
            .ParsedContent = $"[���� CSV �ļ�ʱ����: {ex.Message}]"
        }
        End Try
    End Function

    ' ��ʽ�� CSV �����ݣ�ʹ����׶�
    Private Function FormatCsvRow(row As String, delimiter As Char) As String
        Dim fields As String() = row.Split(delimiter)
        Dim formattedRow As New StringBuilder()

        For i As Integer = 0 To fields.Length - 1
            Dim field As String = fields(i).Trim(""""c) ' �Ƴ�����
            If i < fields.Length - 1 Then
                formattedRow.Append($"{field} | ")
            Else
                formattedRow.Append(field)
            End If
        Next

        Return formattedRow.ToString()
    End Function

    ' ��� CSV �ļ��ķָ���
    Private Function DetectCsvDelimiter(sampleLine As String) As Char
        ' ������ CSV �ָ���
        Dim possibleDelimiters As Char() = {","c, ";"c, vbTab, "|"c}
        Dim bestDelimiter As Char = ","c ' Ĭ��ʹ�ö���
        Dim maxCount As Integer = 0

        ' ���ÿ�����ܵķָ���
        For Each delimiter In possibleDelimiters
            Dim count As Integer = sampleLine.Count(Function(c) c = delimiter)
            If count > maxCount Then
                maxCount = count
                bestDelimiter = delimiter
            End If
        Next

        Return bestDelimiter
    End Function

    ' ����ļ�����
    Private Function DetectFileEncoding(filePath As String) As Encoding
        ' ���ȣ����ǳ��Դ� BOM (Byte Order Mark) ������
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                ' ��ȡǰ�����ֽ������ BOM
                Dim bom(3) As Byte
                Dim bytesRead As Integer = fs.Read(bom, 0, bom.Length)

                ' ����Ƿ��� BOM
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

            ' ����Unicode�滻�ַ������ڼ����Ч�ַ�
            Dim unicodeReplacementChar As Char = ChrW(&HFFFD) ' Unicode �滻�ַ� U+FFFD

            ' ��������ļ������ȳ��� GB18030/GBK ����
            Dim fileExtension As String = Path.GetExtension(filePath).ToLower()
            If fileExtension = ".csv" Then
                ' ���ȳ��� GB18030/GBK ���룬�������Ļ����·ǳ�����
                Try
                    ' ��ȡ�����ļ�����
                    Dim csvSampleBytes As Byte() = New Byte(4095) {}  ' ��ȡǰ 4KB
                    Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                        fs.Read(csvSampleBytes, 0, csvSampleBytes.Length)
                    End Using

                    ' ������ GB18030 ����
                    Dim gbkEncoding As Encoding = Encoding.GetEncoding("GB18030")
                    Dim gbkText As String = gbkEncoding.GetString(csvSampleBytes)

                    ' ���������ı��Ƿ���� CSV ��ʽ���������������źͻ��з���
                    If gbkText.Contains(",") AndAlso (gbkText.Contains(vbCr) OrElse gbkText.Contains(vbLf)) Then
                        ' ����������źͻ��з�����������Ч�� CSV
                        Dim invalidCharCount As Integer = gbkText.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)
                        Dim totalCharCount As Integer = gbkText.Length

                        ' ��������������ʶ���ַ�
                        If invalidCharCount <= totalCharCount * 0.05 Then ' ����5%�Ĳ���ʶ���ַ�
                            Return gbkEncoding
                        End If
                    End If
                Catch ex As Exception
                    ' ���Դ��󣬼���������������
                    Debug.WriteLine($"���� GB18030 ����ʱ����: {ex.Message}")
                End Try
            End If

            ' ���Լ��ֳ����ı���
            Dim encodingsToTry As Encoding() = {
            New UTF8Encoding(False),        ' UTF-8 without BOM
            Encoding.GetEncoding("GB18030"), ' ���ı��룬���Ǽ�������
            Encoding.Default                ' ϵͳĬ�ϱ���
        }

            ' ��ȡ�ļ���ǰ��������
            Dim generalSampleBytes As Byte() = New Byte(4095) {}  ' ��ȡǰ 4KB
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                fs.Read(generalSampleBytes, 0, generalSampleBytes.Length)
            End Using

            Dim bestEncoding As Encoding = encodingsToTry(0) ' Ĭ��ʹ�õ�һ������
            Dim leastInvalidCharCount As Integer = Integer.MaxValue

            ' ����ÿ�ֱ��룬ѡ�����������Ч�ַ��ı���
            For Each enc In encodingsToTry
                Try
                    Dim sample As String = enc.GetString(generalSampleBytes)
                    ' �����ʺź��滻�ַ���������Ϊ��Ч�ַ���ָ��
                    Dim invalidCharCount As Integer = sample.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)

                    ' �����������������Ч�ַ�����
                    If invalidCharCount < leastInvalidCharCount Then
                        leastInvalidCharCount = invalidCharCount
                        bestEncoding = enc

                        ' ���û����Ч�ַ�������ʹ���������
                        If invalidCharCount = 0 Then
                            Exit For
                        End If
                    End If
                Catch ex As Exception
                    ' ���Խ�����󣬳�����һ������
                    Continue For
                End Try
            Next

            ' ʹ�ò���������Ч�ַ��ı���
            Return bestEncoding

        Catch ex As Exception
            Debug.WriteLine($"����ļ�����ʱ����: {ex.Message}")
            ' ����ʱʹ��ϵͳĬ�ϱ���
            Return Encoding.Default
        End Try
    End Function

    Protected MustOverride Function GetApplication() As ApplicationInfo
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCode(vbaCode As String)
    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' ִ�д���ķ���
    Private Sub ExecuteCode(code As String, language As String)
        ' ������������ִ�в�ͬ�Ĳ���
        Select Case language.ToLower()
            Case "vba", "vb", "vbscript", "language-vba", "language-vbscript", "language-vba hljs language-vbscript", "vba hljs language-vbscript"
                ' ִ�� VBA ����
                'ExecuteVBACode(code)
                RunCode(code)
            Case Else
                'MessageBox.Show("��֧�ֵ���������: " & language)
                GlobalStatusStrip.ShowWarning("��֧�ֵ���������: " & language)
        End Select
    End Sub




    ' �������Ƿ������������
    Public Function ContainsProcedureDeclaration(code As String) As Boolean
        ' ʹ�ü򵥵�������ʽ����Ƿ���� Sub �� Function ����
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function


    ' ����ģ���еĵ�һ��������
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
            ' �����������ʹ��������ʽ�Ӵ�������ȡ
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
            GlobalStatusStrip.ShowWarning("�������ô�ģ��ApiKey��")
            Return
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            GlobalStatusStrip.ShowWarning("�������ô�ģ��Api��")
            Return
        End If

        If String.IsNullOrWhiteSpace(question) Then
            GlobalStatusStrip.ShowWarning("���������⣡")
            Return
        End If

        Dim uuid As String = Guid.NewGuid().ToString()

        Try
            Dim requestBody As String = CreateRequestBody(question)
            Await SendHttpRequestStream(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)
            Await SaveFullWebPageAsync2()
        Catch ex As Exception
            MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try

    End Function

    Private Sub ManageHistoryMessageSize()
        ' �����ʷ��Ϣ���������ƣ���һ��system������+1
        While historyMessageData.Count > contextLimit + 1
            ' ����ϵͳ��Ϣ����һ����Ϣ��
            If historyMessageData.Count > 1 Then
                ' �Ƴ��ڶ�����Ϣ������ķ�ϵͳ��Ϣ��
                historyMessageData.RemoveAt(1)
            End If
        End While
    End Sub

    Private Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")

        ' ���� messages ����
        Dim messages As New List(Of String)()

        ' ��� system ��Ϣ
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

        ' ������ʷ��Ϣ��С
        ManageHistoryMessageSize()

        ' �����ʷ��Ϣ
        For Each message In historyMessageData
            messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{message.content}""}}")
        Next

        ' ���� JSON ������
        Dim messagesJson = String.Join(",", messages)
        Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""messages"": [{messagesJson}], ""stream"": true}}"

        Return requestBody
    End Function


    ' ���һ���ṹ���洢token��Ϣ
    Private Structure TokenInfo
        Public PromptTokens As Integer
        Public CompletionTokens As Integer
        Public TotalTokens As Integer
    End Structure

    Private totalTokens As Integer = 0
    Private lastTokenInfo As Nullable(Of TokenInfo)
    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String) As Task

        ' ��װai��ͷ��
        Dim uuid As String = Guid.NewGuid().ToString()
        Try

            ' ǿ��ʹ�� TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = Timeout.InfiniteTimeSpan

                ' ׼������ ---
                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                ' ��ӡ������־ ---
                Debug.WriteLine("[HTTP] ��ʼ������ʽ����...")
                Debug.WriteLine($"[HTTP] Request Body: {requestBody}")


                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                ' �������� ---
                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()
                    Debug.WriteLine($"[HTTP] ��Ӧ״̬��: {response.StatusCode}")

                    Dim js As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{uuid}');"
                    If ChatBrowser.InvokeRequired Then
                        ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(js))
                    Else
                        Await ChatBrowser.ExecuteScriptAsync(js)
                    End If

                    ' ������ ---
                    Dim stringBuilder As New StringBuilder()
                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim buffer(102300) As Char
                            Dim readCount As Integer
                            Do
                                ' ����Ƿ���Ҫֹͣ��ȡ
                                If stopReaderStream Then
                                    Debug.WriteLine("[Stream] �û��ֶ�ֹͣ����ȡ")
                                    ' ��յ�ǰ������
                                    _currentMarkdownBuffer.Clear()
                                    allMarkdownBuffer.Clear()
                                    ' ֹͣ��ȡ���˳�ѭ��
                                    Exit Do
                                End If
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunk As String = New String(buffer, 0, readCount)
                                ' ���chunk������data��ͷ��������
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)
                                'Debug.WriteLine($"[Stream] ���յ���:{stringBuilder}")
                                ' �ж�stringBuilder�Ƿ���'}'��β������������
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), uuid)
                                    stringBuilder.Clear()
                                End If
                            Loop
                            Debug.WriteLine("[Stream] ���������")
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] ��������г���: {ex.ToString()}")
            MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ʹ�����һ����Ӧ���е�total_tokens
            Dim finalTokens As Integer = If(lastTokenInfo.HasValue, lastTokenInfo.Value.TotalTokens, 0)
            Debug.WriteLine($"finally {finalTokens}")
            ExecuteJavaScriptAsyncJS($"processStreamComplete('{uuid}',{finalTokens});")

            ' ��¼ȫ���������У��������ʹ��
            Dim answer = New HistoryMessage() With {
                .role = "assistant",
                .content = allMarkdownBuffer.ToString()
            }
            historyMessageData.Add(answer)
            ' ������ʷ��Ϣ��С
            ManageHistoryMessageSize()

            allMarkdownBuffer.Clear()
            ' ����token��Ϣ
            lastTokenInfo = Nothing
        End Try
    End Function



    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' ���ñ�񡢴�������չ
    .Build()                        ' �������ɱ�ܵ�

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()



    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String)
        Try
            Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                line = line.Trim()
                If line = "[DONE]" Then
                    FlushBuffer("content", uuid) ' ���ˢ�»�����
                    Return
                End If
                If line = "" Then
                    Continue For
                End If

                Debug.Print(line)
                Dim jsonObj As JObject = JObject.Parse(line)

                ' ��ȡtoken��Ϣ - ֻ�������һ����Ӧ���usage��Ϣ
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
                    ' ����Ƿ񵽴�������Ȼ�ָ�㣨���绻�з���
                    'If reasoning_content.Contains(vbLf) OrElse reasoning_content.TrimEnd().EndsWith("`") Then
                    FlushBuffer("reasoning", uuid)
                    'End If
                End If

                Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()

                If Not String.IsNullOrEmpty(content) Then
                    _currentMarkdownBuffer.Append(content)
                    ' ����Ƿ񵽴�������Ȼ�ָ�㣨���绻�з���
                    'If content.Contains(vbLf) OrElse content.TrimEnd().EndsWith("`") Then
                    FlushBuffer("content", uuid)
                    'End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] ���ݴ���ʧ��: {ex.Message}")
            MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


    ' ִ��js�ű����첽����
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
        .Replace("</script>", "<\/script>")  ' ����ű�ע��
    End Function



    ' ���õ�HTTP���󷽷�
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
            MessageBox.Show($"����ʧ��: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' ���ر���HTML�ļ�
    Public Async Function LoadLocalHtmlFile() As Task
        Try
            ' ���HTML�ļ��Ƿ����
            Dim htmlFilePath As String = ChatHtmlFilePath
            If File.Exists(htmlFilePath) Then

                Await Task.Run(Sub()
                                   Dim htmlContent As String = File.ReadAllText(htmlFilePath, System.Text.Encoding.UTF8)
                                   htmlContent = htmlContent.TrimStart("""").TrimEnd("""")
                                   ' ֱ�ӵ���������HTML�ļ�
                                   If ChatBrowser.InvokeRequired Then
                                       ChatBrowser.Invoke(Sub() ChatBrowser.CoreWebView2.NavigateToString(htmlContent))
                                   Else
                                       ChatBrowser.CoreWebView2.NavigateToString(htmlContent)
                                   End If
                               End Sub)

            End If
        Catch ex As Exception
            Debug.WriteLine($"���ر���HTML�ļ�ʱ����{ex.Message}")
        End Try
    End Function

    Public Async Function SaveFullWebPageAsync2() As Task
        Try
            ' 1. ����Ŀ¼��ͬ�������������첽��

            Dim dir = Path.GetDirectoryName(ChatHtmlFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' 2. ��ȡ HTML���첽��������
            Dim htmlContent As String = Await GetFullHtmlContentAsync()

            ' 3. �����ļ����첽��̨�̣߳�
            Await Task.Run(Sub()
                               Dim fullHtml As String = "<!DOCTYPE html>" & Environment.NewLine & htmlContent
                               File.WriteAllText(
                $"{ChatHtmlFilePath}",
                HttpUtility.HtmlDecode(fullHtml),
                System.Text.Encoding.UTF8
            )
                           End Sub)

            Debug.WriteLine("����ɹ�")
        Catch ex As Exception
            Debug.WriteLine($"����ʧ��: {ex.Message}")
        End Try
    End Function

    Private Async Function GetFullHtmlContentAsync() As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' ǿ���л��� WebView2 �� UI �̲߳���
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

                                        ' �Ƴ� <script> ��ǩ��������
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
        ' ����ת���ַ���ֱ�Ӵ� JSON �ַ�������ȡ��
        Return System.Text.RegularExpressions.Regex.Unescape(
        htmlContent
    )
    End Function

    ' ��ʾ�����ã�ÿ�ν���ʹ��1����
    Public Class HistoryMessage
        Public Property role As String
        Public Property content As String
    End Class

    ' ע�븨���ű�
    Protected Sub InitializeWebView2Script()
        ' ���� Web ��Ϣ������
        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived

        ' ��鱾��HTML�ļ��Ƿ���� ���ر���HTML�ļ�
        'Dim htmlFilePath As String = ChatHtmlFilePath
        'If File.Exists(htmlFilePath) Then
        '    LoadLocalHtmlFile()
        'End If

        ' ע�븨���ű�
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

    ' ѡ�����ݷ��͵�������
    Public Async Sub AddSelectedContentItem(sheetName As String, address As String)
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub


    Protected Shared Sub VBAxceptionHandle(ex As Runtime.InteropServices.COMException)
        ' ������������Ȩ������
        If ex.Message.Contains("������ʲ�������") OrElse
       ex.Message.Contains("Programmatic access to Visual Basic Project is not trusted") Then
            VBATrustShowBox()
        Else
            MessageBox.Show("ִ�� VBA ����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Shared Sub VBATrustShowBox()
        MessageBox.Show(
                        "�޷�ִ�� VBA ���룬�밴���²������ã�" & vbCrLf & vbCrLf &
                        "1. ��� '�ļ�' -> 'ѡ��' -> '��������'" & vbCrLf &
                        "2. ��� '������������'" & vbCrLf &
                        "3. ѡ�� '������'" & vbCrLf &
                        "4. ��ѡ '���ζ� VBA ��Ŀ����ģ�͵ķ���'",
                        "��Ҫ������������Ȩ��",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
    End Sub

End Class