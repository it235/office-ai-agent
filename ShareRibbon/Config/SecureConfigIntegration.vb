Imports System.IO
Imports ShareRibbon.Security

''' <summary>
''' 安全配置管理器集成层
''' 将 API Key 从明文 JSON 迁移到加密存储
''' </summary>
Public Class SecureConfigIntegration

    ''' <summary>
    ''' 迁移现有配置到加密存储
    ''' </summary>
    Public Shared Sub MigrateToSecureStorage()
        Try
            ' 读取现有配置
            Dim configPath As String = GetOldConfigPath()
            If Not File.Exists(configPath) Then
                Return
            End If

            Dim json As String = File.ReadAllText(configPath)
            Dim configData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of ConfigManager.ConfigItem))(json)

            ' 迁移每个配置的 API Key
            For Each item In configData
                If Not String.IsNullOrEmpty(item.key) Then
                    ' 保存到加密存储
                    SecureConfigManager.SaveApiKey(item.key, item.platform)

                    ' 清除明文 key（但保留配置项）
                    item.key = "[encrypted]"

                    System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigIntegration] 已迁移 {0} 的 API Key", item.platform))
                End If
            Next

            ' 保存更新后的配置（不含明文 key）
            Dim updatedJson As String = Newtonsoft.Json.JsonConvert.SerializeObject(configData, Newtonsoft.Json.Formatting.Indented)
            File.WriteAllText(configPath, updatedJson)

            System.Diagnostics.Debug.WriteLine("[SecureConfigIntegration] 配置迁移完成")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigIntegration] 迁移失败: {0}", ex.Message))
        End Try
    End Sub

    ''' <summary>
    ''' 加载配置时从加密存储读取 API Key
    ''' </summary>
    Public Shared Sub LoadSecureApiKeys(configData As List(Of ConfigManager.ConfigItem))
        Try
            For Each item In configData
                ' 如果 key 是 [encrypted] 或为空，从加密存储读取
                If String.IsNullOrEmpty(item.key) OrElse item.key = "[encrypted]" Then
                    Dim apiKey As String = SecureConfigManager.LoadApiKey(item.platform)
                    If Not String.IsNullOrEmpty(apiKey) Then
                        item.key = apiKey
                        System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigIntegration] 已加载 {0} 的加密 API Key", item.platform))
                    Else
                        ' 占位符不是 Key：加密存储里没有这台服务商时，内存里必须回到空，
                        ' 否则配置闸门以为已配置，实际把 "[encrypted]" 当 Bearer 发出去。
                        item.key = ""
                    End If
                End If
            Next

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(String.Format("[SecureConfigIntegration] 加载加密 API Key 失败: {0}", ex.Message))
        End Try
    End Sub

    ''' <summary>
    ''' 获取旧配置文件路径
    ''' </summary>
    Private Shared Function GetOldConfigPath() As String
        Dim configFileName As String = "office_ai_config.json"
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            configFileName)
    End Function

    ''' <summary>
    ''' 检查是否需要迁移
    ''' </summary>
    Public Shared Function NeedsMigration() As Boolean
        Try
            Dim configPath As String = GetOldConfigPath()
            If Not File.Exists(configPath) Then
                Return False
            End If

            Dim json As String = File.ReadAllText(configPath)

            ' 检查是否包含明文 key（不是 [encrypted]）
            Return json.Contains("""key"":") AndAlso Not json.Contains("[encrypted]")

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取配置的掩码显示（用于 UI）
    ''' </summary>
    Public Shared Function GetMaskedApiKey(platform As String) As String
        Return SecureConfigManager.GetMaskedApiKey(platform)
    End Function

    ''' <summary>
    ''' 验证 API Key 格式
    ''' </summary>
    Public Shared Function ValidateApiKey(apiKey As String, platform As String) As Boolean
        Return SecureConfigManager.ValidateApiKey(apiKey, platform)
    End Function

End Class
