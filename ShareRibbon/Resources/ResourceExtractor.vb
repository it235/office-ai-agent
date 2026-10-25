Imports System.IO
Imports System.Diagnostics
Imports System.Text

Public Class ResourceExtractor
    Private Shared _lastError As String = String.Empty

    ''' <summary>
    ''' 获取最后一次错误信息
    ''' </summary>
    Public Shared ReadOnly Property LastError As String
        Get
            Return _lastError
        End Get
    End Property

    Public Shared Function ExtractResources() As String
        _lastError = String.Empty

#If DEBUG Then
        DebugListResources()
#End If

        Try
            ' 获取用户本地应用程序数据目录
            Dim appDataPath As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAI",
                "www"
            )

            ' 确保目录存在
            Directory.CreateDirectory(appDataPath)
            Directory.CreateDirectory(Path.Combine(appDataPath, "css"))
            Directory.CreateDirectory(Path.Combine(appDataPath, "js"))

            ' 获取资源管理器 - 使用当前类型的程序集，确保在子插件中也能正确获取 ShareRibbon 资源
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", GetType(ResourceExtractor).Assembly)

            ' 第三方库资源文件映射
            Dim libraryResources As New Dictionary(Of String, String) From {
                {"marked_min", "marked.min.js"},
                {"highlight_min", "highlight.min.js"},
                {"vbscript_min", "vbscript.min.js"},
                {"github_min", "github.min.css"}
            }

            ' 释放第三方库资源
            Dim extractErrors As New List(Of String)()
            Dim refreshed As Integer = 0
            For Each kvp In libraryResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm, refreshed:=refreshed)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 自定义CSS资源文件映射
            Dim cssResources As New Dictionary(Of String, String) From {
                {"styles", "styles.css"}
            }

            ' 释放CSS资源
            For Each kvp In cssResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "css"), rm:=rm, refreshed:=refreshed)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 自定义JS资源文件映射
            Dim jsResources As New Dictionary(Of String, String) From {
                {"office_ai_bridge", "office-ai-bridge.js"},
                {"utils", "utils.js"},
                {"core", "core.js"},
                {"markdown_renderer", "markdown-renderer.js"},
                {"chat_manager", "chat-manager.js"},
                {"message_sender", "message-sender.js"},
                {"code_handler", "code-handler.js"},
                {"settings_manager", "settings-manager.js"},
                {"mcp_manager", "mcp-manager.js"},
                {"revision_manager", "revision-manager.js"},
                {"history_manager", "history-manager.js"},
                {"autocomplete", "autocomplete.js"},
                {"agent_protocol", "agent-protocol.js"},
                {"agent_card", "agent-card.js"},
                {"model_switcher", "model-switcher.js"},
                {"reformat_template", "reformat-template.js"},
                {"reformat_chat", "reformat-chat.js"},
                {"proofread_ui", "proofread-ui.js"},
                {"onboarding", "onboarding.js"},
                {"command_palette", "command-palette.js"}
            }

            ' 释放JS资源
            For Each kvp In jsResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "js"), rm:=rm, refreshed:=refreshed)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next
            
            ' 如果有提取错误，记录但仍然返回路径（部分资源可能已成功）
            If extractErrors.Count > 0 Then
                _lastError = String.Join(Environment.NewLine, extractErrors)
                Debug.WriteLine($"资源提取部分失败: {_lastError}")
            End If

            ' 面板加载的是这里的文件，不是 DLL 里的嵌入资源：不写这行就没人知道真机用的是哪一版前端。
            If refreshed > 0 Then
                AppLogger.Info("ResourceExtractor", $"refreshed {refreshed} static file(s) under {appDataPath}")
            End If

            Return appDataPath
        Catch ex As Exception
            _lastError = $"释放资源失败: {ex.Message}"
            Debug.WriteLine(_lastError)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' 按内容比对释放单个资源：只有字节不同才写盘。曾经用 .version 做整体跳过，
    ''' 结果是改了 JS/CSS 却没人记得 bump 版本，真机一直在跑几个月前的旧前端且毫无报错。
    ''' </summary>
    Private Shared Function ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager, ByRef refreshed As Integer) As String
        Try
            ' 从资源管理器中获取资源
            Dim resourceObj = rm.GetObject(resourceName)

            If resourceObj IsNot Nothing Then
                Dim targetPath = Path.Combine(targetDir, targetFileName)

                ' 统一成字节再比对，避免编码差异导致每次都判定为变化
                Dim data() As Byte
                If TypeOf resourceObj Is Byte() Then
                    data = DirectCast(resourceObj, Byte())
                ElseIf TypeOf resourceObj Is String Then
                    data = New UTF8Encoding(False).GetBytes(DirectCast(resourceObj, String))
                Else
                    Dim errMsg = $"Unsupported resource type for {resourceName}: {resourceObj.GetType().Name}"
                    Debug.WriteLine(errMsg)
                    Return errMsg
                End If

                If File.Exists(targetPath) AndAlso ByteArrayEquals(File.ReadAllBytes(targetPath), data) Then
                    Return String.Empty
                End If

                File.WriteAllBytes(targetPath, data)
                refreshed += 1
                Debug.WriteLine($"Successfully extracted {resourceName} to {targetPath}")
                Return String.Empty
            Else
                Dim errMsg = $"Resource not found: {resourceName}"
                Debug.WriteLine(errMsg)

                ' 列出所有可用的资源，以便调试
                Try
                    Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
                    For Each entry As DictionaryEntry In resourceSet
                        Debug.WriteLine($"Available resource: {entry.Key}")
                    Next
                Catch
                End Try
                
                Return errMsg
            End If
        Catch ex As Exception
            Dim errMsg = $"提取资源 {resourceName} 失败: {ex.Message}"
            Debug.WriteLine(errMsg)
            Return errMsg
        End Try
    End Function

    Private Shared Function ByteArrayEquals(a() As Byte, b() As Byte) As Boolean
        If a Is Nothing OrElse b Is Nothing Then Return False
        If a.Length <> b.Length Then Return False
        For i As Integer = 0 To a.Length - 1
            If a(i) <> b(i) Then Return False
        Next
        Return True
    End Function

    ' 调试方法 - 列出所有资源
    Private Shared Sub DebugListResources()
        Try
            Dim assembly = GetType(ResourceExtractor).Assembly
            Debug.WriteLine("=== 所有嵌入资源 ===")
            For Each resourceName In assembly.GetManifestResourceNames()
                Debug.WriteLine($"Found resource: {resourceName}")
            Next

            Debug.WriteLine("=== Resources 中的资源 ===")
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", assembly)
            Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
            For Each entry As DictionaryEntry In resourceSet
                Debug.WriteLine($"Resource: {entry.Key} (Type: {If(entry.Value IsNot Nothing, entry.Value.GetType().ToString(), "null")})")
            Next
        Catch ex As Exception
            Debug.WriteLine($"列出资源时出错: {ex.Message}")
        End Try
    End Sub
End Class
