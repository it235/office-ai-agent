Imports System.IO
Imports System.Diagnostics

Public Class ResourceExtractor
    Private Shared _lastError As String = String.Empty

    ''' <summary>
    ''' 资源版本号 — 更新此值可强制刷新所有前端资源文件
    ''' </summary>
    Private Shared _resourceVersion As String = "2026.04.28.2"

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

            ' 检查版本标记文件，匹配则跳过提取
            Dim versionFile = Path.Combine(appDataPath, ".version")
            If File.Exists(versionFile) Then
                Try
                    Dim savedVersion = File.ReadAllText(versionFile).Trim()
                    If savedVersion = _resourceVersion Then
                        Debug.WriteLine("[ResourceExtractor] 资源版本匹配，跳过提取")
                        Return appDataPath
                    End If
                Catch
                    ' 读取失败，继续执行提取
                End Try
            End If

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
            For Each kvp In libraryResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm)
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
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "css"), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 自定义JS资源文件映射
            Dim jsResources As New Dictionary(Of String, String) From {
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
                {"intent_preview", "intent-preview.js"},
                {"ralph_loop", "ralph-loop.js"},
                {"ralph_agent", "ralph-agent.js"},
                {"agent_protocol", "agent-protocol.js"},
                {"agent_card", "agent-card.js"},
                {"model_switcher", "model-switcher.js"},
                {"reformat_template", "reformat-template.js"},
                {"config_panel", "config-panel.js"}
            }

            ' 释放JS资源
            For Each kvp In jsResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "js"), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next
            
            ' 如果有提取错误，记录但仍然返回路径（部分资源可能已成功）
            If extractErrors.Count > 0 Then
                _lastError = String.Join(Environment.NewLine, extractErrors)
                Debug.WriteLine($"资源提取部分失败: {_lastError}")
            End If

            ' 写入版本标记文件，下次启动时跳过提取
            Try
                Directory.CreateDirectory(appDataPath)
                File.WriteAllText(Path.Combine(appDataPath, ".version"), _resourceVersion)
            Catch
                ' 写入失败不影响功能
            End Try

            Return appDataPath
        Catch ex As Exception
            _lastError = $"释放资源失败: {ex.Message}"
            Debug.WriteLine(_lastError)
            Return String.Empty
        End Try
    End Function

    Private Shared Function ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager) As String
        Try
            ' 从资源管理器中获取资源
            Dim resourceObj = rm.GetObject(resourceName)

            If resourceObj IsNot Nothing Then
                Dim targetPath = Path.Combine(targetDir, targetFileName)

                ' 根据资源类型处理
                If TypeOf resourceObj Is Byte() Then
                    ' 如果是字节数组，直接写入
                    File.WriteAllBytes(targetPath, DirectCast(resourceObj, Byte()))
                ElseIf TypeOf resourceObj Is String Then
                    ' 如果是字符串，转换为字节后写入
                    File.WriteAllText(targetPath, DirectCast(resourceObj, String))
                Else
                    Dim errMsg = $"Unsupported resource type for {resourceName}: {resourceObj.GetType().Name}"
                    Debug.WriteLine(errMsg)
                    Return errMsg
                End If

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
