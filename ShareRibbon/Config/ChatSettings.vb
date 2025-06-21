Imports System.IO
Imports Newtonsoft.Json

Public Class ChatSettings
    Private ReadOnly _applicationInfo As ApplicationInfo
    Public Sub New(applicationInfo As ApplicationInfo)
        _applicationInfo = applicationInfo
        LoadSettings()
    End Sub

    Public Shared Property topicRandomness As Double = 0.8  ' Ĭ��ֵ��Ϊ Double
    Public Shared Property contextLimit As Integer = 5     ' Ĭ��ֵ��Ϊ Integer
    Public Shared Property selectedCellChecked As Boolean = False
    Public Shared Property executecodePreviewChecked As Boolean = True ' ִ�д���ǰԤ����Ĭ��ѡ��
    Public Shared Property settingsScrollChecked As Boolean = True
    Public Shared Property chatMode As String = "chat"

    ' �޸ķ���ǩ�����������͸�Ϊ Double �� Integer
    Public Sub SaveSettings(topicRandomness As Double, contextLimit As Integer,
                          selectedCell As Boolean, settingsScroll As Boolean, executecodePreview As Boolean, chatMode As String)
        Try
            ' �������ö���
            Dim settings As New Dictionary(Of String, Object) From {
                {"topicRandomness", topicRandomness},
                {"contextLimit", contextLimit},
                {"selectedCellChecked", selectedCell},
                {"settingsScrollChecked", settingsScroll},
                {"executecodePreviewChecked", executecodePreview},
                {"chatMode", chatMode}
            }

            ' �����ñ��浽JSON�ļ�
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            ' ȷ��Ŀ¼����
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath))

            ' ���������л�ΪJSON������
            File.WriteAllText(settingsPath, JsonConvert.SerializeObject(settings, Formatting.Indented))

            ' ���¾�̬����
            ChatSettings.topicRandomness = topicRandomness
            ChatSettings.contextLimit = contextLimit
            ChatSettings.selectedCellChecked = selectedCell
            ChatSettings.settingsScrollChecked = settingsScroll
            ChatSettings.executecodePreviewChecked = executecodePreview
            ChatSettings.chatMode = chatMode

        Catch ex As Exception
            Debug.WriteLine($"��������ʧ��: {ex.Message}")
        End Try
    End Sub

    ' ��������ʱ��������ת��
    Public Sub LoadSettings()
        Try
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            If File.Exists(settingsPath) Then
                ' ��ȡJSON�ļ�
                Dim json = File.ReadAllText(settingsPath)
                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(json)

                ' ���¾�̬���ԣ��������ת��
                If settings.ContainsKey("topicRandomness") Then
                    topicRandomness = Convert.ToDouble(settings("topicRandomness"))
                End If
                If settings.ContainsKey("contextLimit") Then
                    contextLimit = Convert.ToInt32(settings("contextLimit"))
                End If
                If settings.ContainsKey("selectedCellChecked") Then
                    selectedCellChecked = CBool(settings("selectedCellChecked"))
                End If
                If settings.ContainsKey("settingsScrollChecked") Then
                    settingsScrollChecked = CBool(settings("settingsScrollChecked"))
                End If
                If settings.ContainsKey("executecodePreviewChecked") Then
                    executecodePreviewChecked = CBool(settings("executecodePreviewChecked"))
                End If
                If settings.ContainsKey("chatMode") Then
                    chatMode = Convert.ToString(settings("chatMode"))
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"����ChatSettingsʧ��: {ex.Message}")
        End Try
    End Sub
End Class