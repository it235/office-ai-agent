Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 翻译输出模式
''' </summary>
Public Enum TranslateOutputMode
    ''' <summary>替换原文</summary>
    Replace = 0
    ''' <summary>沉浸式翻译（原文+译文并行）</summary>
    Immersive = 1
    ''' <summary>仅显示在侧栏</summary>
    SidePanel = 2
    ''' <summary>新建文档/幻灯片</summary>
    NewDocument = 3
End Enum

Public Class TranslateSettings
    Public Property Enabled As Boolean = False
    Public Property SourceLanguage As String = "auto"
    Public Property TargetLanguage As String = "zh"
    Public Property MaxRequestsPerSecond As Integer = 5
    Public Property EnableSelectionTranslate As Boolean = False
    Public Property PromptText As String = "你是一个专业的翻译，按要求翻译并保留格式。"

    ''' <summary>当前选中的翻译领域</summary>
    Public Property CurrentDomain As String = "通用"

    ''' <summary>翻译输出模式</summary>
    Public Property OutputMode As TranslateOutputMode = TranslateOutputMode.SidePanel

    ''' <summary>沉浸式翻译样式：译文颜色</summary>
    Public Property ImmersiveTranslationColor As String = "#666666"

    ''' <summary>沉浸式翻译样式：是否斜体</summary>
    Public Property ImmersiveTranslationItalic As Boolean = True

    ''' <summary>沉浸式翻译样式：字号比例（相对于原文）</summary>
    Public Property ImmersiveTranslationFontScale As Double = 0.9

    ''' <summary>是否保留原文格式</summary>
    Public Property PreserveFormatting As Boolean = True

    ''' <summary>是否显示翻译进度</summary>
    Public Property ShowProgress As Boolean = True

    ''' <summary>批量翻译时每批段落数（0表示整批翻译，不分批）</summary>
    Public Property BatchSize As Integer = 0

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