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
                'htmlContent = htmlContent.Replace("{wwwroot}", wwwRoot.Replace("\", "/"))

                ' �޸�HTMLģ���е���Դ����
                'Dim htmlContent As String = My.Resources.chat_template
                'htmlContent = htmlContent.Replace(
                '    "href=""css/",
                '    "href=""//officeai.local/css/"
                ').Replace(
                '    "src=""js/",
                '    "src=""//officeai.local/js/"
                ')

                ' ���� HTML ģ��
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)

                ' ���� Marked �ʹ������
                ConfigureMarked()
            Else
                MessageBox.Show("WebView2 ��ʼ��ʧ�ܣ�CoreWebView2 �����á�", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"��ʼ��ʧ��: {ex.Message}{Environment.NewLine}����: {ex.GetType().Name}{Environment.NewLine}��ջ:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    'Protected Async Function InitializeWebView2() As Task
    '    Try
    '        ' ��� WebView2 �Ƿ��Ѿ���ʼ��
    '        If ChatBrowser.CoreWebView2 IsNot Nothing Then
    '            Debug.WriteLine("WebView2 �Ѿ���ʼ����������������")
    '            Return
    '        End If

    '        Debug.WriteLine("��ʼ��ʼ�� WebView2...")

    '        ' �Զ����û�����Ŀ¼
    '        Dim userDataFolder As String = Path.Combine(
    '        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    '        "MyAppWebView2Cache"
    '    )

    '        ' ȷ��Ŀ¼����
    '        If Not Directory.Exists(userDataFolder) Then
    '            Directory.CreateDirectory(userDataFolder)
    '        End If

    '        ' ��UI�߳������� CreationProperties
    '        If ChatBrowser.InvokeRequired Then
    '            Await ChatBrowser.Invoke(Sub()
    '                                         ' �����µ� Environment ����
    '                                         Dim envOptions = New CoreWebView2EnvironmentOptions()
    '                                         ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
    '                .UserDataFolder = userDataFolder
    '            }
    '                                     End Sub)
    '        Else
    '            ' �����µ� Environment ����
    '            Dim envOptions = New CoreWebView2EnvironmentOptions()
    '            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
    '            .UserDataFolder = userDataFolder
    '        }
    '        End If

    '        Debug.WriteLine("���ڳ�ʼ�� CoreWebView2...")

    '        ' ȷ����UI�߳��ϳ�ʼ�� WebView2
    '        If ChatBrowser.InvokeRequired Then
    '            Await ChatBrowser.Invoke(Async Function()
    '                                         Await ChatBrowser.EnsureCoreWebView2Async(Nothing)
    '                                     End Function)
    '        Else
    '            Await ChatBrowser.EnsureCoreWebView2Async(Nothing)
    '        End If

    '        ' ȷ�� CoreWebView2 �ѳ�ʼ��
    '        If ChatBrowser.CoreWebView2 IsNot Nothing Then
    '            Debug.WriteLine("CoreWebView2 ��ʼ���ɹ������ڼ���ģ��...")

    '            ' ���� HTML ģ��
    '            If ChatBrowser.InvokeRequired Then
    '                ChatBrowser.Invoke(Sub() ChatBrowser.CoreWebView2.NavigateToString(My.Resources.chat_template))
    '            Else
    '                ChatBrowser.CoreWebView2.NavigateToString(My.Resources.chat_template)
    '            End If

    '            ' ���� Marked �ʹ������
    '            Await ConfigureMarked()

    '            Debug.WriteLine("ģ��������")
    '        Else
    '            Throw New Exception("WebView2 ��ʼ��ʧ�ܣ�CoreWebView2 �����á�")
    '        End If

    '    Catch ex As Exception
    '        Debug.WriteLine($"WebView2��ʼ��ʧ��: {ex.Message}")
    '        Debug.WriteLine($"��ջ����: {ex.StackTrace}")
    '        Throw
    '    End Try
    'End Function

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

    '        ' ����HTMLģ��
    '        Await LoadLocalHtmlFile()
    '    Catch ex As Exception
    '        MessageBox.Show($"WebView2��ʼ��ʧ��: {ex.Message}")
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


    ' ִ�д���ķ���
    Private Sub ExecuteCode(code As String, language As String)
        ' ������������ִ�в�ͬ�Ĳ���
        Select Case language.ToLower()
            Case "vba", "vb", "vbscript", "language-vba", "language-vbscript", "language-vba hljs language-vbscript", "vba hljs language-vbscript"
                ' ִ�� VBA ����
                ExecuteVBACode(code)
            Case Else
                'MessageBox.Show("��֧�ֵ���������: " & language)
                GlobalStatusStrip.ShowWarning("��֧�ֵ���������: " & language)
        End Select
    End Sub


    ' ִ��ǰ�˴����� VBA ����Ƭ��
    Private Sub ExecuteVBACode(vbaCode As String)
        ' ��ȡ VBA ��Ŀ
        Dim vbProj As VBProject = GetVBProject()

        ' ��ӿ�ֵ���
        If vbProj Is Nothing Then
            Return
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' ������ʱģ��
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' �������Ƿ��Ѱ��� Sub/Function ����
            If ContainsProcedureDeclaration(vbaCode) Then
                ' �����Ѱ�������������ֱ�����
                vbComp.CodeModule.AddFromString(vbaCode)

                ' ���ҵ�һ����������ִ��
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    RunCode(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("�޷��ڴ������ҵ���ִ�еĹ���")
                    GlobalStatusStrip.ShowWarning("�޷��ڴ������ҵ���ִ�еĹ���")
                End If
            Else
                ' ���벻�������������������װ�� Auto_Run ������
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' ִ�� Auto_Run ����
                RunCode(tempModuleName & ".Auto_Run")
            End If

        Catch ex As Exception
            MessageBox.Show("ִ�� VBA ����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ���۳ɹ�����ʧ�ܣ���ɾ����ʱģ��
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' �����������
            End Try
        End Try
    End Sub


    ' �������Ƿ������������
    Private Function ContainsProcedureDeclaration(code As String) As Boolean
        ' ʹ�ü򵥵�������ʽ����Ƿ���� Sub �� Function ����
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function


    ' ����ģ���еĵ�һ��������
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

        ' �����ʷ��Ϣ
        For Each message In historyMessageData
            messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{message.content}""}}")
        Next

        ' ���� JSON ������
        Dim messagesJson = String.Join(",", messages)
        Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""messages"": [{messagesJson}], ""stream"": true}}"

        Return requestBody
        ' ʹ�ô� ConfigSettings �л�ȡ��ģ������
        'Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""system"", ""content"": """ & ConfigSettings.propmtContent & """},{""role"": ""user"", ""content"": """ & result & """}],""stream"":true}"
    End Function


    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String) As Task
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

                ' ��װai��ͷ��
                Dim uuid As String = Guid.NewGuid().ToString()

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
                            Dim buffer(10230) As Char
                            Dim readCount As Integer
                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunk As String = New String(buffer, 0, readCount)
                                ' ���chunk������data��ͷ��������
                                'If Not chunk.StartsWith("data:") Then Continue Do
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)
                                'Debug.WriteLine($"[Stream] ���յ���:{stringBuilder}")
                                ' �ж�stringBuilder�Ƿ���'}'��β������������
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), uuid)
                                    stringBuilder.Clear()
                                End If

                                'If Not line.StartsWith("{") OrElse Not line.EndsWith("}") Then
                                '    _currentMarkdownBuffer.Append(line)
                                '    Continue For
                                'End If

                            Loop
                            Debug.WriteLine("[Stream] ���������")
                            Await ExecuteJavaScriptAsyncJS($"processStreamComplete('{uuid}');")
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] ��������г���: {ex.ToString()}")
            MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ��¼ȫ���������У��������ʹ��
            Dim answer = New HistoryMessage() With {
                .role = "assistant",
                .content = allMarkdownBuffer.ToString()
            }
            historyMessageData.Add(answer)
            allMarkdownBuffer.Clear()
        End Try
    End Function



    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' ���ñ�񡢴�������չ
    .Build()                        ' �������ɱ�ܵ�

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()

    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String)
        Try
            'Dim lines As String() = rawChunk.Split({"data:"}, StringSplitOptions.RemoveEmptyEntries)
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

        ' ��鱾��HTML�ļ��Ƿ����
        Dim htmlFilePath As String = ChatHtmlFilePath
        If File.Exists(htmlFilePath) Then
            ' ���ر���HTML�ļ�
            LoadLocalHtmlFile()
        End If

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

    ' ѡ�����ݷ��͵�������
    Protected Async Sub AddSelectedContentItem(sheetName As String, address As String)
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