Imports System.IO

Public Class ResourceExtractor
    Public Shared Function ExtractResources() As String
        DebugListResources()
        Try
            ' 获取用户本地应用数据目录
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

            ' 资源名到文件名的映射
            Dim resources As New Dictionary(Of String, String) From {
                {"marked_min", "marked.min.js"},
                {"highlight_min", "highlight.min.js"},
                {"vbscript_min", "vbscript.min.js"},
                {"github_min", "github.min.css"}
            }

            ' 释放资源
            For Each kvp In resources
                ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm)
            Next

            Return appDataPath
        Catch ex As Exception
            Debug.WriteLine($"释放资源失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Shared Sub ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager)
        Try
            ' 从资源管理器获取资源
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

                ' 输出所有可用的资源名称以便调试
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