'Imports System.IO
'Imports System.Reflection
'Imports System.Runtime.InteropServices
'Imports Microsoft.Win32

'' COM接口定义
'<ComVisible(True)>
'<Guid("12345678-1234-1234-1234-123456789012")> ' 生成一个新的GUID
'<InterfaceType(ComInterfaceType.InterfaceIsDual)>
'Public Interface IExcelAiFunctions
'    Function TLLM(prompt As String) As String
'    Function CLLM(prompt As String, Optional model As String = "",
'                             Optional systemPrompt As String = "",
'                             Optional temperature As Double = 0.7,
'                             Optional maxTokens As Integer = 1000) As String
'End Interface

'' 实现接口的COM类
'<ComVisible(True)>
'<Guid("87654321-4321-4321-4321-210987654321")> ' 生成一个新的GUID
'<ProgId("ExcelAi.Functions")>
'<ClassInterface(ClassInterfaceType.None)>
'Public Class ExcelAiFunctions
'    Implements IExcelAiFunctions

'    Private excelFunctions As New ExcelFunctions()

'    Public Function TLLM(prompt As String) As String Implements IExcelAiFunctions.TLLM
'        Try
'            Return excelFunctions.TLLM(prompt)
'        Catch ex As Exception
'            Return $"错误: {ex.Message}"
'        End Try
'    End Function

'    Public Function CLLM(prompt As String, Optional model As String = "",
'                                    Optional systemPrompt As String = "",
'                                    Optional temperature As Double = 0.7,
'                                    Optional maxTokens As Integer = 1000) As String Implements IExcelAiFunctions.CLLM
'        Try
'            Return excelFunctions.CLLM(prompt, model, systemPrompt, temperature, maxTokens)
'        Catch ex As Exception
'            Return $"错误: {ex.Message}"
'        End Try
'    End Function

'    ' 手动注册COM组件的辅助方法
'    Public Shared Sub RegisterFunction()
'        Try
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("ExcelAi.Functions")
'            regKey.SetValue("", "Excel AI Functions")

'            ' 添加CLSID项
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            clsidKey.SetValue("", "{87654321-4321-4321-4321-210987654321}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("COM ProgID已手动注册")
'        Catch ex As Exception
'            ' 如果无法访问注册表，则记录错误
'            System.Diagnostics.Debug.WriteLine($"注册函数时发生错误: {ex.Message}")
'        End Try
'    End Sub

'    ' 提供COM注册和注销的辅助类
'    <ComVisible(False)>
'    Public Class ComRegistrationHelper
'        ' 注册COM组件
'        Public Shared Sub RegisterCom()
'            Try
'                ' 获取当前程序集路径
'                Dim assemblyPath As String = Assembly.GetExecutingAssembly().Location

'                ' 使用regasm.exe注册COM组件
'                Dim regasmPath As String = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "regasm.exe")
'                Dim process As New System.Diagnostics.Process()
'                process.StartInfo.FileName = regasmPath
'                process.StartInfo.Arguments = $"/codebase ""{assemblyPath}"""
'                process.StartInfo.UseShellExecute = True
'                process.StartInfo.Verb = "runas"  ' 请求管理员权限
'                process.StartInfo.CreateNoWindow = False
'                process.Start()
'                process.WaitForExit()

'                If process.ExitCode = 0 Then
'                    System.Diagnostics.Debug.WriteLine("COM组件注册成功")
'                Else
'                    System.Diagnostics.Debug.WriteLine($"COM组件注册失败，退出代码: {process.ExitCode}")
'                End If
'            Catch ex As Exception
'                System.Diagnostics.Debug.WriteLine($"注册COM组件时出错: {ex.Message}")
'            End Try
'        End Sub
'    End Class

'    ' 添加COM注册方法
'    <ComRegisterFunction()>
'    Public Shared Sub RegisterFunction(ByVal type As Type)
'        Try
'            System.Diagnostics.Debug.WriteLine($"COM注册: {type.Name}")

'            ' 注册到 CurrentUser\Software\Classes (这是 CurrentUser 的正确位置)
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\ExcelAi.Functions")
'            regKey.SetValue("", "Excel AI Functions")

'            ' 添加CLSID项
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            clsidKey.SetValue("", "{87654321-4321-4321-4321-210987654321}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("COM注册成功")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"COM注册失败: {ex.Message}")
'        End Try
'    End Sub

'    ' 添加COM注销方法
'    <ComUnregisterFunction()>
'    Public Shared Sub UnregisterFunction(ByVal type As Type)
'        Try
'            Registry.CurrentUser.DeleteSubKeyTree("Software\Classes\ExcelAi.Functions", False)
'            System.Diagnostics.Debug.WriteLine("COM注销成功")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"COM注销失败: {ex.Message}")
'        End Try
'    End Sub
'End Class