Imports System.IO
Imports ShareRibbon.Security


Public Class ConfigManager
    Public Shared Property ConfigData As List(Of ConfigItem)

    ' 默认配置文件在当前用户，我的文档下
    Private Shared configFileName As String = "office_ai_config.json"
    Private Shared configFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder, configFileName)

    Public Sub LoadConfig()
        ' 初始化配置数据
        ConfigData = New List(Of ConfigItem)()

        ' 添加默认配置
        If Not File.Exists(configFilePath) Then
            ' 首次使用，加载预置配置
            ConfigData.AddRange(PresetProviders.GetCloudProviders())
            ConfigData.AddRange(PresetProviders.GetLocalProviders())
        Else
            ' 加载用户已有配置
            Dim json As String = File.ReadAllText(configFilePath)
            Dim loadedData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of ConfigItem))(json)

            ' 合并预置配置和用户配置
            MergeConfigurations(loadedData)
        End If

        ' 检查是否需要迁移到加密存储
        If SecureConfigIntegration.NeedsMigration() Then
            System.Diagnostics.Debug.WriteLine("[ConfigManager] 检测到明文 API Key，开始迁移到加密存储")
            SecureConfigIntegration.MigrateToSecureStorage()
        End If

        ' 从加密存储加载 API Key
        SecureConfigIntegration.LoadSecureApiKeys(ConfigData)

        ApplyConfigSnapshot()
    End Sub

    ''' <summary>
    ''' 把"当前选中的服务商 + 选中的聊天模型"投射到全局快照 ConfigSettings。
    ''' 必须无条件赋值：没有选中项、或选中项的 URL/Key 被用户清空时，也要把旧值抹掉。
    ''' 否则保存配置后面板仍按上一次的 URL 和 Key 发请求，未配置闸门也会被旧 Key 放行。
    ''' EmbeddingModel 与提示词字段不由本方法管理，它们有各自的写入方，重置会打断补全和向量检索。
    ''' </summary>
    Public Shared Sub ApplyConfigSnapshot()
        Dim chosen = If(ConfigData Is Nothing,
                        Nothing,
                        ConfigData.FirstOrDefault(Function(item) item.selected))

        ConfigSettings.platform = If(chosen?.platform, "")
        ConfigSettings.ApiUrl = If(chosen?.url, "")
        ConfigSettings.ApiKey = If(chosen?.key, "")
        ConfigSettings.ModelName = ""
        ConfigSettings.mcpable = False
        ConfigSettings.ReasoningMode = "default"
        ConfigSettings.fimSupported = False
        ConfigSettings.fimUrl = ""

        If chosen Is Nothing Then Return

        Dim chatModel = chosen.model?.FirstOrDefault(
            Function(m) m.selected AndAlso m.modelType = ModelType.Chat)
        If chatModel Is Nothing Then Return

        ConfigSettings.ModelName = chatModel.modelName
        ConfigSettings.mcpable = chatModel.mcpable
        ConfigSettings.ReasoningMode = chatModel.reasoningMode
        ConfigSettings.fimSupported = chatModel.fimSupported
        ConfigSettings.fimUrl = If(String.IsNullOrEmpty(chatModel.fimUrl), chosen.url, chatModel.fimUrl)
    End Sub

    ''' <summary>
    ''' 合并预置配置和用户配置
    ''' 保留用户的key、selected、validated等状态
    ''' </summary>
    Private Sub MergeConfigurations(loadedData As List(Of ConfigItem))
        ' 先添加云端预置配置
        For Each preset In PresetProviders.GetCloudProviders()
            Dim existing = loadedData.FirstOrDefault(Function(x) x.platform = preset.platform OrElse x.url = preset.url)
            If existing IsNot Nothing Then
                ' 保留用户的key和selected状态
                preset.key = existing.key
                preset.selected = existing.selected
                preset.validated = existing.validated
                preset.translateSelected = existing.translateSelected
                ' 合并模型列表
                For Each userModel In existing.model
                    Dim presetModel = preset.model.FirstOrDefault(Function(x) x.modelName = userModel.modelName)
                    If presetModel IsNot Nothing Then
                        presetModel.selected = userModel.selected
                        presetModel.mcpValidated = userModel.mcpValidated
                        presetModel.mcpable = userModel.mcpable
                        presetModel.translateSelected = userModel.translateSelected
                        presetModel.fimSupported = userModel.fimSupported
                        presetModel.fimUrl = userModel.fimUrl
                        presetModel.reasoningMode = userModel.reasoningMode
                    Else
                        ' 用户添加的自定义模型，保留
                        preset.model.Add(userModel)
                    End If
                Next
            End If
            ConfigData.Add(preset)
        Next

        ' 添加本地预置配置
        For Each preset In PresetProviders.GetLocalProviders()
            Dim existing = loadedData.FirstOrDefault(Function(x) x.platform = preset.platform OrElse x.url = preset.url)
            If existing IsNot Nothing Then
                preset.key = existing.key
                preset.selected = existing.selected
                preset.validated = existing.validated
                preset.translateSelected = existing.translateSelected
                ' 本地模型URL可能被用户修改
                If Not String.IsNullOrEmpty(existing.url) Then
                    preset.url = existing.url
                End If
                ' 合并模型列表
                For Each userModel In existing.model
                    Dim presetModel = preset.model.FirstOrDefault(Function(x) x.modelName = userModel.modelName)
                    If presetModel IsNot Nothing Then
                        presetModel.selected = userModel.selected
                        presetModel.mcpValidated = userModel.mcpValidated
                        presetModel.mcpable = userModel.mcpable
                        presetModel.translateSelected = userModel.translateSelected
                        presetModel.fimSupported = userModel.fimSupported
                        presetModel.fimUrl = userModel.fimUrl
                        presetModel.reasoningMode = userModel.reasoningMode
                    Else
                        preset.model.Add(userModel)
                    End If
                Next
            End If
            ConfigData.Add(preset)
        Next

        ' 添加用户自定义的非预置配置
        For Each userConfig In loadedData
            Dim isPresetPlatform = ConfigData.Any(Function(x) x.platform = userConfig.platform OrElse x.url = userConfig.url)
            If Not isPresetPlatform Then
                ' 用户自定义的服务商，直接添加
                userConfig.isPreset = False
                ConfigData.Add(userConfig)
            End If
        Next
    End Sub


    ' 保存到文件中，默认存在用户的文档目录下
    Public Shared Sub SaveConfig()
        ' API Key 只进 DPAPI（secure.config），配置文件里只写占位符：
        ' 明文密钥落在"我的文档"等于向任何能读该目录的进程公开。
        ' 只有确认加密写入成功的项才剥掉明文，写盘失败的保持原值，不能让用户的 Key 凭空消失。
        Dim originals As New List(Of String)()
        Dim hidden As New List(Of Boolean)()
        If ConfigData IsNot Nothing Then
            For Each item In ConfigData
                originals.Add(If(item.key, ""))
                Dim canHide As Boolean = False
                If Not String.IsNullOrEmpty(item.key) AndAlso item.key <> "[encrypted]" Then
                    canHide = SecureConfigManager.SaveApiKey(item.key, item.platform)
                End If
                hidden.Add(canHide)
                If canHide Then item.key = "[encrypted]"
            Next
        End If

        Try
            Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(ConfigData, Newtonsoft.Json.Formatting.Indented)
            ' 如果configFilePath的目录不存在就创建
            Dim dir = Path.GetDirectoryName(configFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            '如果文件不存在就创建
            If Not File.Exists(configFilePath) Then
                File.Create(configFilePath).Dispose()
            End If
            File.WriteAllText(configFilePath, json)
        Finally
            ' 内存里的 ConfigData 必须继续拿着真实 Key：MCP、意图识别、翻译都直接读 cfg.key，
            ' 它们只在启动时经过 LoadConfig，保存后立刻清空会让这些路径拿到占位符。
            If ConfigData IsNot Nothing Then
                For i As Integer = 0 To ConfigData.Count - 1
                    If hidden(i) Then ConfigData(i).key = originals(i)
                Next
            End If
        End Try

        ' 快照必须跟着磁盘上的选择走，否则配置窗口清空 URL/Key 后，
        ' 面板仍会用保存前的旧值发请求，E3-03 的未配置闸门也会被旧 Key 放行。
        ApplyConfigSnapshot()
    End Sub


    ' Api配置（每次仅可使用1格）
    Public Class ConfigItem
        Public Property platform As String
        Public Property url As String
        Public Property model As List(Of ConfigItemModel)
        Public Property key As String
        Public Property selected As Boolean

        ' 是否被选为翻译专用平台（在 UI 中为单选，仅允许一个 true）
        Public Property translateSelected As Boolean = False

        ' 是否通过了API验证
        Public Property validated As Boolean

        ' 服务商类型: Cloud(云端) / Local(本地)
        Public Property providerType As ProviderType = ProviderType.Cloud

        ' 获取APIKey的注册链接
        Public Property registerUrl As String = ""

        ' 是否为预置配置
        Public Property isPreset As Boolean = False

        ' 本地模型默认APIKey提示
        Public Property defaultApiKey As String = ""

        Public Overrides Function ToString() As String
            Return platform
        End Function
    End Class

    ' 具体模型，例：阿里云百炼的 qwen-coder-plus
    Public Class ConfigItemModel
        Public Property modelName As String
        Public Property selected As Boolean

        ' 是否被选为翻译专用平台（在 UI 中为单选，仅允许一个 true）
        Public Property translateSelected As Boolean = False
        Public Property mcpable As Boolean = False
        Public Property mcpValidated As Boolean = False
        
        ' FIM (Fill-In-the-Middle) 补全能力支持
        Public Property fimSupported As Boolean = False
        
        ' FIM API端点（如果与chat端点不同）
        Public Property fimUrl As String = ""

        ' 是否为推理模型
        Public Property isReasoningModel As Boolean = False

        ' 推理开关: default=不传参数, enabled=显式开启, disabled=显式关闭
        Public Property reasoningMode As String = "default"

        ' 显示名称(含[推理][MCP]标签)
        Public Property displayName As String = ""
        
        ' 模型类型：Chat(对话) / Embedding(向量)
        Public Property modelType As ModelType = ModelType.Chat
        
        Public Overrides Function ToString() As String
            Return If(String.IsNullOrEmpty(displayName), modelName, displayName)
        End Function
    End Class
    
    ' 模型类型枚举
    Public Enum ModelType
        Chat = 0
        Embedding = 1
    End Enum
End Class
