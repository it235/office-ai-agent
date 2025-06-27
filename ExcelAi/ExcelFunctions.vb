'Imports System.Diagnostics
'Imports System.IO
'Imports System.Net
'Imports System.Net.Http
'Imports System.Reflection
'Imports System.Runtime.InteropServices
'Imports System.Text
'Imports System.Threading.Tasks
'Imports Microsoft.Office.Interop.Excel
'Imports Microsoft.Win32
'Imports Newtonsoft.Json
'Imports Newtonsoft.Json.Linq
'Imports ShareRibbon

'' ������Excel��ע��Ͷ����Զ��庯��
'<ComVisible(True)>
'<ClassInterface(ClassInterfaceType.AutoDual)>
'<ProgId("ExcelAi.ExcelFunctions")>
'Public Class ExcelFunctions

'    ' LLM�ı����ɺ��� - �����汾
'    ' ������prompt - ��ʾ��
'    <ComVisible(True)>
'    Public Function TLLM(prompt As String) As String
'        Try
'            ' ʹ��Ĭ�����õ��������溯��
'            Return CLLM(prompt, "", "", 0.7, 1000)
'        Catch ex As Exception
'            Return $"����: {ex.Message}"
'        End Try
'    End Function

'    ' LLM�ı����ɺ��� - �߼��汾
'    ' ������
'    ' - prompt: ��ʾ��
'    ' - model: ģ������ (��ѡ)
'    ' - systemPrompt: ϵͳ��ʾ�� (��ѡ)
'    ' - temperature: �¶Ȳ��� (��ѡ)
'    ' - maxTokens: ������������� (��ѡ)

'    <ComVisible(True)>
'    Public Function CLLM(
'        prompt As String,
'        Optional model As String = "",
'        Optional systemPrompt As String = "",
'        Optional temperature As Double = 0.7,
'        Optional maxTokens As Integer = 1000) As String

'        Try
'            ' ��֤����
'            If String.IsNullOrEmpty(prompt) Then
'                Return "����: ��ʾ�ʲ���Ϊ��"
'            End If

'            ' ʹ��Ĭ��ֵ�����δ�ṩ��
'            Dim apiKey As String = GetApiKey()
'            Dim apiUrl As String = GetApiUrl()

'            If String.IsNullOrEmpty(apiKey) Then
'                Return "����: δ����API��Կ"
'            End If

'            If String.IsNullOrEmpty(apiUrl) Then
'                Return "����: δ����API URL"
'            End If

'            ' ʹ��ָ����ģ�ͻ�Ĭ��ģ��
'            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

'            ' ����������
'            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)

'            ' ����API����ȡ�����ʹ����ʽ����
'            Dim response As String = LLMUtil.SendHttpRequest(apiUrl, apiKey, requestBody).Result

'            ' �����ӦΪ�գ����ش�����Ϣ
'            If String.IsNullOrEmpty(response) Then
'                Return "����: APIδ������Ӧ"
'            End If
'            Dim parsedResponse As JObject = JObject.Parse(response)
'            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()
'            Return cellValue
'        Catch ex As Exception
'            Return $"����: {ex.Message}"
'        End Try
'    End Function


'    ' ��ȡAPI��Կ
'    Private Function GetApiKey() As String
'        Try
'            ' �����ù�������ȡAPI��Կ
'            Return ShareRibbon.ConfigSettings.ApiKey
'        Catch ex As Exception
'            Return ""
'        End Try
'    End Function

'    ' ��ȡAPI URL
'    Private Function GetApiUrl() As String
'        Try
'            ' �����ù�������ȡAPI URL
'            Return ShareRibbon.ConfigSettings.ApiUrl
'        Catch ex As Exception
'            Return ""
'        End Try
'    End Function

'    ' ��ȡĬ��ģ��
'    Private Function GetDefaultModel() As String
'        Try
'            ' �����ù�������ȡĬ��ģ��
'            Dim model As String = ShareRibbon.ConfigSettings.ModelName

'            ' ���δ���ã�ʹ��ͨ��Ĭ��ֵ
'            If String.IsNullOrEmpty(model) Then
'                Return "gpt-3.5-turbo"
'            End If

'            Return model
'        Catch ex As Exception
'            Return "gpt-3.5-turbo"
'        End Try
'    End Function

'    ' ���COMע�᷽��
'    <ComRegisterFunction()>
'    Public Shared Sub RegisterFunction(ByVal type As Type)
'        Try
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COMע��: {type.Name}")

'            ' ���ע�����
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey($"ExcelAi.ExcelFunctions")
'            regKey.SetValue("", "Excel AI Functions Implementation")

'            ' ���CLSID��
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            ' ��ȡ���͵�GUID
'            Dim guidAttr As GuidAttribute = CType(type.GetCustomAttributes(GetType(GuidAttribute), False)(0), GuidAttribute)
'            clsidKey.SetValue("", $"{{{guidAttr.Value}}}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("ExcelFunctions COMע��ɹ�")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COMע��ʧ��: {ex.Message}")
'        End Try
'    End Sub

'    ' ���COMע������
'    <ComUnregisterFunction()>
'    Public Shared Sub UnregisterFunction(ByVal type As Type)
'        Try
'            Registry.CurrentUser.DeleteSubKeyTree($"ExcelAi.ExcelFunctions", False)
'            System.Diagnostics.Debug.WriteLine("ExcelFunctions COMע���ɹ�")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COMע��ʧ��: {ex.Message}")
'        End Try
'    End Sub
'End Class