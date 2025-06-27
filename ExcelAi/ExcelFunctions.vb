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

'' 用于在Excel中注册和定义自定义函数
'<ComVisible(True)>
'<ClassInterface(ClassInterfaceType.AutoDual)>
'<ProgId("ExcelAi.ExcelFunctions")>
'Public Class ExcelFunctions

'    ' LLM文本生成函数 - 基本版本
'    ' 参数：prompt - 提示词
'    <ComVisible(True)>
'    Public Function TLLM(prompt As String) As String
'        Try
'            ' 使用默认设置调用完整版函数
'            Return CLLM(prompt, "", "", 0.7, 1000)
'        Catch ex As Exception
'            Return $"错误: {ex.Message}"
'        End Try
'    End Function

'    ' LLM文本生成函数 - 高级版本
'    ' 参数：
'    ' - prompt: 提示词
'    ' - model: 模型名称 (可选)
'    ' - systemPrompt: 系统提示词 (可选)
'    ' - temperature: 温度参数 (可选)
'    ' - maxTokens: 最大生成令牌数 (可选)

'    <ComVisible(True)>
'    Public Function CLLM(
'        prompt As String,
'        Optional model As String = "",
'        Optional systemPrompt As String = "",
'        Optional temperature As Double = 0.7,
'        Optional maxTokens As Integer = 1000) As String

'        Try
'            ' 验证输入
'            If String.IsNullOrEmpty(prompt) Then
'                Return "错误: 提示词不能为空"
'            End If

'            ' 使用默认值（如果未提供）
'            Dim apiKey As String = GetApiKey()
'            Dim apiUrl As String = GetApiUrl()

'            If String.IsNullOrEmpty(apiKey) Then
'                Return "错误: 未配置API密钥"
'            End If

'            If String.IsNullOrEmpty(apiUrl) Then
'                Return "错误: 未配置API URL"
'            End If

'            ' 使用指定的模型或默认模型
'            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

'            ' 创建请求体
'            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)

'            ' 调用API并获取结果（使用流式处理）
'            Dim response As String = LLMUtil.SendHttpRequest(apiUrl, apiKey, requestBody).Result

'            ' 如果响应为空，返回错误信息
'            If String.IsNullOrEmpty(response) Then
'                Return "错误: API未返回响应"
'            End If
'            Dim parsedResponse As JObject = JObject.Parse(response)
'            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()
'            Return cellValue
'        Catch ex As Exception
'            Return $"错误: {ex.Message}"
'        End Try
'    End Function


'    ' 获取API密钥
'    Private Function GetApiKey() As String
'        Try
'            ' 从配置管理器获取API密钥
'            Return ShareRibbon.ConfigSettings.ApiKey
'        Catch ex As Exception
'            Return ""
'        End Try
'    End Function

'    ' 获取API URL
'    Private Function GetApiUrl() As String
'        Try
'            ' 从配置管理器获取API URL
'            Return ShareRibbon.ConfigSettings.ApiUrl
'        Catch ex As Exception
'            Return ""
'        End Try
'    End Function

'    ' 获取默认模型
'    Private Function GetDefaultModel() As String
'        Try
'            ' 从配置管理器获取默认模型
'            Dim model As String = ShareRibbon.ConfigSettings.ModelName

'            ' 如果未配置，使用通用默认值
'            If String.IsNullOrEmpty(model) Then
'                Return "gpt-3.5-turbo"
'            End If

'            Return model
'        Catch ex As Exception
'            Return "gpt-3.5-turbo"
'        End Try
'    End Function

'    ' 添加COM注册方法
'    <ComRegisterFunction()>
'    Public Shared Sub RegisterFunction(ByVal type As Type)
'        Try
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COM注册: {type.Name}")

'            ' 添加注册表项
'            Dim regKey As RegistryKey = Registry.CurrentUser.CreateSubKey($"ExcelAi.ExcelFunctions")
'            regKey.SetValue("", "Excel AI Functions Implementation")

'            ' 添加CLSID项
'            Dim clsidKey As RegistryKey = regKey.CreateSubKey("CLSID")
'            ' 获取类型的GUID
'            Dim guidAttr As GuidAttribute = CType(type.GetCustomAttributes(GetType(GuidAttribute), False)(0), GuidAttribute)
'            clsidKey.SetValue("", $"{{{guidAttr.Value}}}")

'            regKey.Close()

'            System.Diagnostics.Debug.WriteLine("ExcelFunctions COM注册成功")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COM注册失败: {ex.Message}")
'        End Try
'    End Sub

'    ' 添加COM注销方法
'    <ComUnregisterFunction()>
'    Public Shared Sub UnregisterFunction(ByVal type As Type)
'        Try
'            Registry.CurrentUser.DeleteSubKeyTree($"ExcelAi.ExcelFunctions", False)
'            System.Diagnostics.Debug.WriteLine("ExcelFunctions COM注销成功")
'        Catch ex As Exception
'            System.Diagnostics.Debug.WriteLine($"ExcelFunctions COM注销失败: {ex.Message}")
'        End Try
'    End Sub
'End Class