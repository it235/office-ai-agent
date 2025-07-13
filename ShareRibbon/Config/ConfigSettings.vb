' 存储配置的api大模型和api key
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property mcpable As Boolean

    ' 提示词相关配置
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"
End Class
