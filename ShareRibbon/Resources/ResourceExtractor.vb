Imports System.IO

Public Class ResourceExtractor
    Public Shared Function ExtractResources() As String
        DebugListResources()
        Try
            ' ��ȡ�û�����Ӧ������Ŀ¼
            Dim appDataPath As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAI",
                "www"
            )

            ' ȷ��Ŀ¼����
            Directory.CreateDirectory(appDataPath)
            Directory.CreateDirectory(Path.Combine(appDataPath, "css"))
            Directory.CreateDirectory(Path.Combine(appDataPath, "js"))

            ' ��ȡ��Դ������
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", System.Reflection.Assembly.GetExecutingAssembly())

            ' ��Դ�����ļ�����ӳ��
            Dim resources As New Dictionary(Of String, String) From {
                {"marked_min", "marked.min.js"},
                {"highlight_min", "highlight.min.js"},
                {"vbscript_min", "vbscript.min.js"},
                {"github_min", "github.min.css"}
            }

            ' �ͷ���Դ
            For Each kvp In resources
                ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm)
            Next

            Return appDataPath
        Catch ex As Exception
            Debug.WriteLine($"�ͷ���Դʧ��: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Shared Sub ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager)
        Try
            ' ����Դ��������ȡ��Դ
            Dim resourceObj = rm.GetObject(resourceName)

            If resourceObj IsNot Nothing Then
                Dim targetPath = Path.Combine(targetDir, targetFileName)

                ' ������Դ���ʹ���
                If TypeOf resourceObj Is Byte() Then
                    ' ������ֽ����飬ֱ��д��
                    File.WriteAllBytes(targetPath, DirectCast(resourceObj, Byte()))
                ElseIf TypeOf resourceObj Is String Then
                    ' ������ַ�����ת��Ϊ�ֽں�д��
                    File.WriteAllText(targetPath, DirectCast(resourceObj, String))
                Else
                    Debug.WriteLine($"Unsupported resource type for {resourceName}: {resourceObj.GetType().Name}")
                    Return
                End If

                Debug.WriteLine($"Successfully extracted {resourceName} to {targetPath}")
            Else
                Debug.WriteLine($"Resource not found: {resourceName}")

                ' ������п��õ���Դ�����Ա����
                Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
                For Each entry As DictionaryEntry In resourceSet
                    Debug.WriteLine($"Available resource: {entry.Key}")
                Next
            End If
        Catch ex As Exception
            Debug.WriteLine($"��ȡ��Դ {resourceName} ʧ��: {ex.Message}")
        End Try
    End Sub

    ' ���Է��� - �г�������Դ
    Private Shared Sub DebugListResources()
        Try
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Debug.WriteLine("=== ����Ƕ����Դ ===")
            For Each resourceName In assembly.GetManifestResourceNames()
                Debug.WriteLine($"Found resource: {resourceName}")
            Next

            Debug.WriteLine("=== Resources �е���Դ ===")
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", assembly)
            Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
            For Each entry As DictionaryEntry In resourceSet
                Debug.WriteLine($"Resource: {entry.Key} (Type: {If(entry.Value IsNot Nothing, entry.Value.GetType().ToString(), "null")})")
            Next
        Catch ex As Exception
            Debug.WriteLine($"�г���Դʱ����: {ex.Message}")
        End Try
    End Sub
End Class