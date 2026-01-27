Imports System.IO

Public Class ResourceExtractor
    Public Shared Function ExtractResources() As String
        DebugListResources()
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

            ' 获取资源管理器
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", System.Reflection.Assembly.GetExecutingAssembly())

            ' 第三方库资源文件映射
            Dim libraryResources As New Dictionary(Of String, String) From {
                {"marked_min", "marked.min.js"},
                {"highlight_min", "highlight.min.js"},
                {"vbscript_min", "vbscript.min.js"},
                {"github_min", "github.min.css"}
            }

            ' 释放第三方库资源
            For Each kvp In libraryResources
                ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm)
            Next

            ' 自定义CSS资源文件映射
            Dim cssResources As New Dictionary(Of String, String) From {
                {"styles", "styles.css"}
            }

            ' 释放CSS资源
            For Each kvp In cssResources
                ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "css"), rm:=rm)
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
                {"intent_preview", "intent-preview.js"}
            }

            ' 释放JS资源
            For Each kvp In jsResources
                ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "js"), rm:=rm)
            Next

            Return appDataPath
        Catch ex As Exception
            Debug.WriteLine($"释放资源失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Shared Sub ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager)
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
                    Debug.WriteLine($"Unsupported resource type for {resourceName}: {resourceObj.GetType().Name}")
                    Return
                End If

                Debug.WriteLine($"Successfully extracted {resourceName} to {targetPath}")
            Else
                Debug.WriteLine($"Resource not found: {resourceName}")

                ' 列出所有可用的资源，以便调试
                Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
                For Each entry As DictionaryEntry In resourceSet
                    Debug.WriteLine($"Available resource: {entry.Key}")
                Next
            End If
        Catch ex As Exception
            Debug.WriteLine($"提取资源 {resourceName} 失败: {ex.Message}")
        End Try
    End Sub

    ' 调试方法 - 列出所有资源
    Private Shared Sub DebugListResources()
        Try
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
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
