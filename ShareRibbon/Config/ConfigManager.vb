Imports System.IO
Imports Newtonsoft.Json

Public Class ConfigManager
    Public Shared Property ConfigData As List(Of ConfigItem)

    ' 默认配置文件在当前用户，我的文档下
    Private Shared configFileName As String = "office_ai_config.json"
    Private Shared configFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder, configFileName)

    Public Sub LoadConfig()
        ' 初始化配置数据
        ConfigData = New List(Of ConfigItem)()

        Dim ollama = New ConfigItem() With {
                .pltform = "Ollama本地模型",
                .url = "http://localhost:11434/v1/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "deepseek-r1:1.5b", .selected = True},
                    New ConfigItemModel() With {.modelName = "deepseek-r1:7b", .selected = False},
                    New ConfigItemModel() With {.modelName = "deepseek-r1:14b", .selected = False}
                },
                .key = "",
                .selected = False,
                .translateSelected = False
            }

        Dim ds = New ConfigItem() With {
                .pltform = "深度求索",
                .url = "https://api.deepseek.com/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "deepseek-chat", .selected = True},
                    New ConfigItemModel() With {.modelName = "deepseek-reasoner", .selected = False}
                },
                .key = "",
                .selected = True,
                .translateSelected = True
            }
        Dim aliyun = New ConfigItem() With {
                .pltform = "阿里云百炼",
                .url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "qwen-coder-plus", .selected = True},
                    New ConfigItemModel() With {.modelName = "qwen-max", .selected = False},
                    New ConfigItemModel() With {.modelName = "qwen-plus", .selected = False}
                },
                .key = "",
                .selected = False,
                .translateSelected = False
            }
        Dim volces = New ConfigItem() With {
            .pltform = "百度千帆",
            .url = "https://qianfan.baidubce.com/v2/chat/completions",
            .model = New List(Of ConfigItemModel) From {
                New ConfigItemModel() With {.modelName = "deepseek-v3", .selected = True},
                New ConfigItemModel() With {.modelName = "deepseek-r1", .selected = False}
            },
            .key = "",
            .selected = False,
            .translateSelected = False
        }
        Dim siliconflow = New ConfigItem() With {
            .pltform = "华为硅基流动",
            .url = "https://api.siliconflow.cn/v1/chat/completions",
            .model = New List(Of ConfigItemModel) From {
            New ConfigItemModel() With {.modelName = "deepseek-ai/DeepSeek-V3", .selected = True},
            New ConfigItemModel() With {.modelName = "deepseek-ai/DeepSeek-R1", .selected = False}
            },
            .key = "",
            .selected = False,
            .translateSelected = False
        }

        ' 添加默认配置
        If Not File.Exists(configFilePath) Then
            ConfigData.Add(ds)
            ConfigData.Add(siliconflow)
            ConfigData.Add(aliyun)
            ConfigData.Add(ollama)
            ConfigData.Add(volces)
        Else
            ' 加载自定义配置
            Dim json As String = File.ReadAllText(configFilePath)
            ConfigData = JsonConvert.DeserializeObject(Of List(Of ConfigItem))(json)
        End If

        ' 初始化配置，将数据初始化到 ConfigSettings，方便全局调用
        For Each item In ConfigData
            If item.selected Then
                ConfigSettings.ApiUrl = item.url
                ConfigSettings.ApiKey = item.key
                ConfigSettings.platform = item.pltform
                For Each item_m In item.model
                    If item_m.selected Then
                        ConfigSettings.ModelName = item_m.modelName
                        ConfigSettings.mcpable = item_m.mcpable
                        ConfigSettings.fimSupported = item_m.fimSupported
                        ConfigSettings.fimUrl = If(String.IsNullOrEmpty(item_m.fimUrl), item.url, item_m.fimUrl)
                    End If
                Next
            End If
        Next
    End Sub


    ' 保存到文件中，默认存在用户的文档目录下
    Public Shared Sub SaveConfig()
        Dim json As String = JsonConvert.SerializeObject(ConfigData, Formatting.Indented)
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

    End Sub


    ' Api配置（每次仅可使用1格）
    Public Class ConfigItem
        Public Property pltform As String
        Public Property url As String
        Public Property model As List(Of ConfigItemModel)
        Public Property key As String
        Public Property selected As Boolean

        ' 是否被选为翻译专用平台（在 UI 中为单选，仅允许一个 true）
        Public Property translateSelected As Boolean = False

        ' 是否通过了API验证
        Public Property validated As Boolean

        Public Overrides Function ToString() As String
            Return pltform
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
        
        Public Overrides Function ToString() As String
            Return modelName
        End Function
    End Class
End Class
