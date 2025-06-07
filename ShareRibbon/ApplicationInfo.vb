' ShareRibbon\Models\ApplicationInfo.vb
Public Class ApplicationInfo
    Public Property Name As String
    Public Property Type As OfficeApplicationType

    Public Sub New(name As String, type As OfficeApplicationType)
        Me.Name = name
        Me.Type = type
    End Sub

    ' 获取提示词配置文件路径
    Public Function GetPromptConfigFilePath() As String
        Dim fileName As String = $"office_ai_prompt_config_{Me.Name.ToLower()}.json"
        Return System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            fileName)
    End Function
End Class

Public Enum OfficeApplicationType
    Word
    Excel
    PowerPoint
End Enum