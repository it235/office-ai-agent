'Imports System.IO
'Imports System.Reflection
'Imports System.Runtime.InteropServices
'Imports Microsoft.Win32

'' COM�ӿڶ���
'<ComVisible(True)>
'<Guid("12345678-1234-1234-1234-123456789012")> ' ����һ���µ�GUID
'<InterfaceType(ComInterfaceType.InterfaceIsDual)>
'Public Interface IExcelAiFunctions
'    Function TLLM(prompt As String) As String
'    Function CLLM(prompt As String, Optional model As String = "",
'                             Optional systemPrompt As String = "",
'                             Optional temperature As Double = 0.7,
'                             Optional maxTokens As Integer = 1000) As String
'End Interface

'' ʵ�ֽӿڵ�COM��
'<ComVisible(True)>
'<Guid("87654321-4321-4321-4321-210987654321")> ' ����һ���µ�GUID
'<ProgId("ExcelAi.Functions")>
'<ClassInterface(ClassInterfaceType.None)>
'Public Class ExcelAiFunctions
'    Implements IExcelAiFunctions

'    Private excelFunctions As New ExcelFunctions()

'    Public Function TLLM(prompt As String) As String Implements IExcelAiFunctions.TLLM
'        Try
'            Return excelFunctions.TLLM(prompt)
'        Catch ex As Exception
'            Return $"����: {ex.Message}"
'        End Try
'    End Function

'    Public Function CLLM(prompt As String, Optional model As String = "",
'                                    Optional systemPrompt As String = "",
'                                    Optional temperature As Double = 0.7,
'                                    Optional maxTokens As Integer = 1000) As String Implements IExcelAiFunctions.CLLM
'        Try
'            Return excelFunctions.CLLM(prompt, model, systemPrompt, temperature, maxTokens)
'        Catch ex As Exception
'            Return $"����: {ex.Message}"
'        End Try
'    End Function

'    ' �ֶ�ע��COM����ĸ�������
'    Public Shared Sub RegisterFunction()
'        Try
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("ExcelAi.Functions")
'            regKey.SetValue("", "Excel AI Functions")

'            ' ���CLSID��
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            clsidKey.SetValue("", "{87654321-4321-4321-4321-210987654321}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("COM ProgID���ֶ�ע��")
'        Catch ex As Exception
'            ' ����޷�����ע������¼����
'            System.Diagnostics.Debug.WriteLine($"ע�ắ��ʱ��������: {ex.Message}")
'        End Try
'    End Sub

'    ' �ṩCOMע���ע���ĸ�����
'    <ComVisible(False)>
'    Public Class ComRegistrationHelper
'        ' ע��COM���
'        Public Shared Sub RegisterCom()
'            Try
'                ' ��ȡ��ǰ����·��
'                Dim assemblyPath As String = Assembly.GetExecutingAssembly().Location

'                ' ʹ��regasm.exeע��COM���
'                Dim regasmPath As String = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "regasm.exe")
'                Dim process As New System.Diagnostics.Process()
'                process.StartInfo.FileName = regasmPath
'                process.StartInfo.Arguments = $"/codebase ""{assemblyPath}"""
'                process.StartInfo.UseShellExecute = True
'                process.StartInfo.Verb = "runas"  ' �������ԱȨ��
'                process.StartInfo.CreateNoWindow = False
'                process.Start()
'                process.WaitForExit()

'                If process.ExitCode = 0 Then
'                    System.Diagnostics.Debug.WriteLine("COM���ע��ɹ�")
'                Else
'                    System.Diagnostics.Debug.WriteLine($"COM���ע��ʧ�ܣ��˳�����: {process.ExitCode}")
'                End If
'            Catch ex As Exception
'                System.Diagnostics.Debug.WriteLine($"ע��COM���ʱ����: {ex.Message}")
'            End Try
'        End Sub
'    End Class

'    ' ���COMע�᷽��
'    <ComRegisterFunction()>
'    Public Shared Sub RegisterFunction(ByVal type As Type)
'        Try
'            System.Diagnostics.Debug.WriteLine($"COMע��: {type.Name}")

'            ' ע�ᵽ CurrentUser\Software\Classes (���� CurrentUser ����ȷλ��)
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\ExcelAi.Functions")
'            regKey.SetValue("", "Excel AI Functions")

'            ' ���CLSID��
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            clsidKey.SetValue("", "{87654321-4321-4321-4321-210987654321}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("COMע��ɹ�")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"COMע��ʧ��: {ex.Message}")
'        End Try
'    End Sub

'    ' ���COMע������
'    <ComUnregisterFunction()>
'    Public Shared Sub UnregisterFunction(ByVal type As Type)
'        Try
'            Registry.CurrentUser.DeleteSubKeyTree("Software\Classes\ExcelAi.Functions", False)
'            System.Diagnostics.Debug.WriteLine("COMע���ɹ�")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"COMע��ʧ��: {ex.Message}")
'        End Try
'    End Sub
'End Class