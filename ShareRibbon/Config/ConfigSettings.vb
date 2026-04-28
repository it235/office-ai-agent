' 存储配置的api大模型和api key
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property mcpable As Boolean

    ' Embedding 模型配置
    Public Shared Property EmbeddingModel As String = ""

    ' FIM (Fill-In-the-Middle) 补全能力
    Public Shared Property fimSupported As Boolean = False
    Public Shared Property fimUrl As String = ""

    ' 提示词相关配置
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    ' Agent 架构切换：true=使用新 AgentKernel，false=保留旧 RalphLoop/RalphAgent
    Public Shared Property UseNewAgentKernel As Boolean = True

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"
End Class
