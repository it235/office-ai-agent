Imports System.IO
Imports Newtonsoft.Json

Public Class ConfigManager
    Public Shared Property ConfigData As List(Of ConfigItem)

    ' Ĭ�������ļ��ڵ�ǰ�û����ҵ��ĵ���
    Private Shared configFileName As String = "office_ai_config.json"
    Private Shared configFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder, configFileName)

    Public Sub LoadConfig()
        ' ��ʼ����������
        ConfigData = New List(Of ConfigItem)()

        Dim ollama = New ConfigItem() With {
                .pltform = "Ollama����ģ��",
                .url = "http://localhost:11434/v1/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "deepseek-r1:1.5b", .selected = True},
                    New ConfigItemModel() With {.modelName = "deepseek-r1:7b", .selected = False},
                    New ConfigItemModel() With {.modelName = "deepseek-r1:14b", .selected = False}
                },
                .key = "",
                .selected = False
            }

        Dim ds = New ConfigItem() With {
                .pltform = "�������",
                .url = "https://api.deepseek.com/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "deepseek-chat", .selected = True},
                    New ConfigItemModel() With {.modelName = "deepseek-reasoner", .selected = False}
                },
                .key = "",
                .selected = True
            }
        Dim aliyun = New ConfigItem() With {
                .pltform = "�����ư���",
                .url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
                .model = New List(Of ConfigItemModel) From {
                    New ConfigItemModel() With {.modelName = "qwen-coder-plus", .selected = True},
                    New ConfigItemModel() With {.modelName = "qwen-max", .selected = False},
                    New ConfigItemModel() With {.modelName = "qwen-plus", .selected = False}
                },
                .key = "",
                .selected = False
            }
        Dim volces = New ConfigItem() With {
            .pltform = "�ٶ�ǧ��",
            .url = "https://qianfan.baidubce.com/v2/chat/completions",
            .model = New List(Of ConfigItemModel) From {
                New ConfigItemModel() With {.modelName = "deepseek-v3", .selected = True},
                New ConfigItemModel() With {.modelName = "deepseek-r1", .selected = False}
            },
            .key = "",
            .selected = False
        }
        Dim siliconflow = New ConfigItem() With {
            .pltform = "��Ϊ�������",
            .url = "https://api.siliconflow.cn/v1/chat/completions",
            .model = New List(Of ConfigItemModel) From {
            New ConfigItemModel() With {.modelName = "deepseek-ai/DeepSeek-V3", .selected = True},
            New ConfigItemModel() With {.modelName = "deepseek-ai/DeepSeek-R1", .selected = False}
            },
            .key = "",
            .selected = False
        }

        ' ���Ĭ������
        If Not File.Exists(configFilePath) Then
            ConfigData.Add(ds)
            ConfigData.Add(siliconflow)
            ConfigData.Add(aliyun)
            ConfigData.Add(ollama)
            ConfigData.Add(volces)
        Else
            ' �����Զ�������
            Dim json As String = File.ReadAllText(configFilePath)
            ConfigData = JsonConvert.DeserializeObject(Of List(Of ConfigItem))(json)
        End If

        ' ��ʼ�����ã������ݳ�ʼ���� ConfigSettings������ȫ�ֵ���
        For Each item In ConfigData
            If item.selected Then
                ConfigSettings.ApiUrl = item.url
                ConfigSettings.ApiKey = item.key
                ConfigSettings.platform = item.pltform
                For Each item_m In item.model
                    If item_m.selected Then
                        ConfigSettings.ModelName = item_m.modelName
                    End If
                Next
            End If
        Next

    End Sub

    ' ���浽�ļ��У�Ĭ�ϴ����û����ĵ�Ŀ¼��
    Public Shared Sub SaveConfig()
        Dim json As String = JsonConvert.SerializeObject(ConfigData, Formatting.Indented)
        ' ���configFilePath��Ŀ¼�����ھʹ���
        Dim dir = Path.GetDirectoryName(configFilePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        '����ļ������ھʹ���
        If Not File.Exists(configFilePath) Then
            File.Create(configFilePath).Dispose()
        End If
        File.WriteAllText(configFilePath, json)

    End Sub


    ' Api���ã�ÿ�ν���ʹ��1��
    Public Class ConfigItem
        Public Property pltform As String
        Public Property url As String
        Public Property model As List(Of ConfigItemModel)
        Public Property key As String
        Public Property selected As Boolean

        ' �Ƿ�ͨ����API��֤
        Public Property validated As Boolean

        Public Overrides Function ToString() As String
            Return pltform
        End Function
    End Class

    ' ����ģ�ͣ����������ư����� qwen-coder-plus
    Public Class ConfigItemModel
        Public Property modelName As String
        Public Property selected As Boolean
        Public Overrides Function ToString() As String
            Return modelName
        End Function
    End Class
End Class
