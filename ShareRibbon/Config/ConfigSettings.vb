' �洢���õ�api��ģ�ͺ�api key
Public Class ConfigSettings
    ' Public NotInheritable Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String

    ' ��ʾ���������
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"
End Class
