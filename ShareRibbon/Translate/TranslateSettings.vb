Imports System.IO
Imports Newtonsoft.Json

Public Class TranslateSettings
    Public Property Enabled As Boolean = False
    Public Property SourceLanguage As String = "auto"
    Public Property TargetLanguage As String = "zh"
    Public Property MaxRequestsPerSecond As Integer = 5
    Public Property EnableSelectionTranslate As Boolean = False
    Public Property PromptText As String = "你是一个专业的翻译，按要求翻译并保留格式。"

    Private Shared ReadOnly fileName As String = "translate_config.json"
    Private Shared ReadOnly filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                                              ConfigSettings.OfficeAiAppDataFolder, fileName)

    Public Shared Function Load() As TranslateSettings
        Try
            If Not File.Exists(filePath) Then
                Dim def As New TranslateSettings()
                def.Save()
                Return def
            End If
            Dim json As String = File.ReadAllText(filePath)
            Return JsonConvert.DeserializeObject(Of TranslateSettings)(json)
        Catch ex As Exception
            Return New TranslateSettings()
        End Try
    End Function

    Public Sub Save()
        Try
            Dim dir = Path.GetDirectoryName(filePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Dim json As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
            File.WriteAllText(filePath, json)
        Catch
            ' 忽略写入错误
        End Try
    End Sub
End Class